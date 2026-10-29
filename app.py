# -*- coding: utf-8 -*-
"""KATIKA DATA LAB : métrés, devis, risques et carnet de chantier.

Prototype pédagogique français. Les transactions sont exclusivement simulées.
Le registre et le dataset CSV sont conservés hors du dossier livré.
"""
from __future__ import annotations
import base64
import copy
from datetime import date, datetime, timezone
from decimal import Decimal, ROUND_HALF_UP
import hashlib
import html
import io
import json
import math
import os
from pathlib import Path
import sqlite3
import uuid
import warnings
from urllib.parse import quote as urlquote

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from PIL import Image, ImageDraw, ImageOps
from scipy import stats
import streamlit as st
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

APP_NAME='KATIKA DATA LAB'
PURPLE,ORANGE,BLUE,YELLOW='#A089FF','#FF8B25','#369BDE','#FFD226'
INK,BG,PANEL='#EFF1FA','#0D1019','#171B28'
METHODS=['Orange Money','MTN MoMo']
STATUSES=['En attente','Réussi','Échoué','Annulé']
MAX_BYTES=25*1024*1024
NAV=['Tableau de bord','Métrés et devis','Risques et données','Encyclopédie BTP','Paiements','Carnet de chantier','Données et aide']


def stamp():return datetime.now(timezone.utc).isoformat(timespec='seconds')
def uid():return uuid.uuid4().hex
def esc(value):return html.escape(str(value))
def fmt(value,d=0):return f'{float(value):,.{d}f}'.replace(',',' ').replace('.',',')
def money(value):return fmt(value)+' FCFA'


def number(value,label,low=0,high=1e12,integer=False):
    try:value=float(str(value).replace(',','.'))
    except (TypeError,ValueError):raise ValueError(label+' : nombre valide attendu.') from None
    if not math.isfinite(value) or not low<=value<=high or (integer and not value.is_integer()):
        raise ValueError(f'{label} : valeur {"entière " if integer else ""}entre {low} et {high} attendue.')
    return int(value) if integer else value


def text(value,label,limit=100,required=False):
    value=str(value or '').strip()
    if (required and not value) or len(value)>limit:raise ValueError(f'{label} : texte {"obligatoire, " if required else ""}{limit} caractères maximum.')
    return value


def day(value):
    try:return date.fromisoformat(str(value))
    except (ValueError,TypeError):raise ValueError('Date attendue au format AAAA-MM-JJ.') from None


def rounded(value):return int(Decimal(str(value)).quantize(Decimal('1'),rounding=ROUND_HALF_UP))


def default_risk():
    return dict(resource_cv=.15,cost_cv=.08,cv_cement=.12,cv_steel=.15,cv_sand=.10,
                rho_cs=.65,rho_csa=.32,rho_ssa=.48,rho_rc=.45,
                alpha=3.8,beta=2.,incident_probability=.3,max_extra=150000.,
                prior_om=1.,prior_momo=1.,fee_om=0.,fee_momo=0.,count=10000,seed=42)


def default_measure():
    return dict(label='Dalle rectangulaire',length=5.,width=4.,height=.15,dosage=350.,bag=50.,
                sand=.5,gravel=.8,water_ratio=.5,steel=80.,waste=5.,labor=100000.,tax=0.,
                price_cement=5500.,price_sand=20000.,price_gravel=25000.,price_water=25.,price_steel=650.)


def validate_measure(raw):
    result={'label':text(raw.get('label'),'Désignation',120,True)}
    bounds={'length':(.001,10000),'width':(.001,10000),'height':(.001,1000),'dosage':(0,1000),
            'bag':(1,100),'sand':(0,3),'gravel':(0,3),'water_ratio':(0,2),'steel':(0,1000),
            'waste':(0,100),'labor':(0,1e10),'tax':(0,100),
            'price_cement':(0,1e8),'price_sand':(0,1e8),'price_gravel':(0,1e8),'price_water':(0,1e6),'price_steel':(0,1e8)}
    for key,(low,high) in bounds.items():result[key]=number(raw.get(key),key,low,high)
    return result


def estimate(raw):
    x=validate_measure(raw);volume=x['length']*x['width']*x['height'];factor=1+x['waste']/100
    theoretical=dict(cement=volume*x['dosage'],sand=volume*x['sand'],gravel=volume*x['gravel'],
                     water=volume*x['dosage']*x['water_ratio'],steel=volume*x['steel'])
    quantities={key:value*factor for key,value in theoretical.items()}
    bags=int(math.ceil(quantities['cement']/x['bag']-1e-12))
    rows=[dict(key='cement',name='Ciment',unit=f'sac de {x["bag"]:g} kg',quantity=bags,price=x['price_cement']),
          dict(key='sand',name='Sable',unit='m³',quantity=quantities['sand'],price=x['price_sand']),
          dict(key='gravel',name='Gravier',unit='m³',quantity=quantities['gravel'],price=x['price_gravel']),
          dict(key='water',name='Eau',unit='L',quantity=quantities['water'],price=x['price_water']),
          dict(key='steel',name='Acier',unit='kg',quantity=quantities['steel'],price=x['price_steel'])]
    for r in rows:r['amount']=rounded(Decimal(str(r['quantity']))*Decimal(str(r['price'])))
    subtotal=sum(r['amount'] for r in rows)+rounded(x['labor']);tax=rounded(subtotal*x['tax']/100)
    total=subtotal+tax
    if total>1e12:raise ValueError('Le devis dépasse la limite de 1 000 milliards de FCFA.')
    return dict(inputs=x,surface=x['length']*x['width'],volume=volume,theoretical=theoretical,
                quantities=quantities,bags=bags,rows=rows,subtotal=subtotal,tax=tax,total=total)


def validate_risk(raw):
    bounds={'resource_cv':(0,1),'cost_cv':(0,1),'cv_cement':(0,1),'cv_steel':(0,1),'cv_sand':(0,1),
            'rho_cs':(-1,1),'rho_csa':(-1,1),'rho_ssa':(-1,1),'rho_rc':(-.99,.99),
            'alpha':(.1,100),'beta':(.1,100),'incident_probability':(0,1),'max_extra':(0,1e10),
            'prior_om':(.1,1000),'prior_momo':(.1,1000),'fee_om':(0,10),'fee_momo':(0,10),
            'count':(1000,100000),'seed':(0,2_000_000_000)}
    out={k:number(raw.get(k),k,*b,integer=k in ('count','seed')) for k,b in bounds.items()}
    if np.linalg.eigvalsh(correlation(out)).min() < -1e-9:raise ValueError('Les trois corrélations ne forment pas une matrice valide. Réduire leurs écarts ou reprendre 0,65 / 0,32 / 0,48.')
    return out


def correlation(r):return np.array([[1,r['rho_cs'],r['rho_csa']],[r['rho_cs'],1,r['rho_ssa']],[r['rho_csa'],r['rho_ssa'],1]],float)


def lognormal_mean(mean,cv,z):
    sigma=np.sqrt(np.log1p(cv*cv))
    return mean*np.exp(sigma*z-.5*sigma*sigma)


def posterior(transactions,city,config):
    rows=[t for t in transactions if t['status']=='Réussi' and t['city'].casefold()==city.casefold()]
    om=sum(t['method']=='Orange Money' for t in rows);momo=len(rows)-om
    a=config['prior_om']+om;b=config['prior_momo']+momo
    return dict(count=len(rows),om=om,momo=momo,empirical=om/len(rows) if rows else None,
                alpha=a,beta=b,mean=a/(a+b),interval=stats.beta.ppf([.025,.975],a,b).tolist())


@st.cache_data(ttl=1800,max_entries=8,show_spinner=False)
def simulate(resource,reference,weights,raw,p_om):
    r=validate_risk(raw);resource=number(resource,'Ressources moyennes');reference=number(reference,'Coût de référence')
    weights=np.asarray(weights,float)
    if weights.shape!=(3,) or not np.isfinite(weights).all() or (weights<0).any() or weights.sum()>1+1e-9:raise ValueError('Pondérations des matériaux invalides.')
    p_om=number(p_om,'Probabilité Orange Money',0,1);n=r['count'];rng=np.random.default_rng(r['seed'])
    vals,vecs=np.linalg.eigh(correlation(r));root=vecs@np.diag(np.sqrt(np.clip(vals,0,None)))
    material_z=rng.standard_normal((n,3))@root.T
    multipliers=np.column_stack([lognormal_mean(1,r[k],material_z[:,i]) for i,k in enumerate(['cv_cement','cv_steel','cv_sand'])])
    z_cost=rng.standard_normal(n);z_resource=r['rho_rc']*z_cost+math.sqrt(1-r['rho_rc']**2)*rng.standard_normal(n)
    resources=lognormal_mean(resource,r['resource_cv'],z_resource)
    base=reference*(1+((multipliers-1)*weights).sum(axis=1))*lognormal_mean(1,r['cost_cv'],z_cost)
    u=rng.beta(r['alpha'],r['beta'],n)
    incident=rng.random(n)<r['incident_probability']
    method=rng.random(n)<p_om
    fees=reference*np.where(method,r['fee_om'],r['fee_momo'])/100
    expenses=base+incident*u*r['max_extra']+fees
    balance=resources-expenses;hits=int(np.count_nonzero(balance<0));p=hits/n
    z=1.95996398454;den=1+z*z/n;mid=(p+z*z/(2*n))/den;half=z*math.sqrt(p*(1-p)/n+z*z/(4*n*n))/den
    pct=(multipliers-1)*100
    std=pct.std(axis=0,ddof=1);cov=np.cov(pct,rowvar=False)
    empirical_corr=np.divide(cov,np.outer(std,std),out=np.full((3,3),np.nan),where=np.outer(std,std)>0)
    return dict(resource=resources,expenses=expenses,balance=balance,probability=p,hits=hits,
                interval=(max(0,mid-half),min(1,mid+half)),beta_index=float(-stats.norm.ppf(p)) if 0<p<1 else None,
                covariance=cov,empirical_correlation=empirical_corr,latent=correlation(r),method_om=method,
                z_resource=z_resource[:1500],z_cost=z_cost[:1500],beta_samples=u[:5000],config=r)


def empty_store():return dict(schema=1,projects=[],quotes=[],transactions=[],observations=[])


def new_code(records,prefix):
    used={r['id'] for r in records};i=1
    while f'{prefix}-{i:04d}' in used:i+=1
    return f'{prefix}-{i:04d}'


def new_project(name,city,budget):
    return dict(id='',name=text(name,'Nom du chantier',100,True),city=text(city,'Ville',80,True),
                budget=number(budget,'Ressources prévues'),progress=0.,latitude=None,longitude=None,
                photo='',created=stamp(),risk=default_risk())


def validate_media(encoded,max_size,label):
    if not isinstance(encoded,str):raise ValueError(label+' invalide.')
    if encoded:
        try:blob=base64.b64decode(encoded,validate=True)
        except Exception:raise ValueError(label+' invalide.') from None
        if len(blob)>max_size:raise ValueError(label+' trop volumineux.')
        if label=='Photo':
            try:
                with Image.open(io.BytesIO(blob)) as image:
                    if image.format!='JPEG' or image.width*image.height>20000000:raise ValueError('Image hors limites.')
                    image.verify()
            except Exception:raise ValueError('Photo sauvegardée illisible.') from None
    return encoded


def validate_store(raw):
    if not isinstance(raw,dict) or raw.get('schema')!=1:raise ValueError('Sauvegarde KATIKA version 1 attendue.')
    data=copy.deepcopy(raw)
    limits={'projects':30,'quotes':300,'transactions':2000,'observations':500}
    for key,limit in limits.items():
        if not isinstance(data.get(key),list) or len(data[key])>limit:raise ValueError(f'{key} : maximum {limit} éléments.')
        if any(not isinstance(x,dict) for x in data[key]) or len({x.get('id') for x in data[key]})!=len(data[key]):raise ValueError('Références dupliquées ou invalides.')
        for x in data[key]:text(x.get('id'),'Référence',50,True)
    project_ids={p['id'] for p in data['projects']}
    for p in data['projects']:
        p['name']=text(p.get('name'),'Nom du chantier',100,True);p['city']=text(p.get('city'),'Ville',80,True)
        p['budget']=number(p.get('budget'),'Budget');p['progress']=number(p.get('progress'),'Avancement',0,100)
        p['risk']=validate_risk(p.get('risk',default_risk()))
        if (p.get('latitude') is None)!=(p.get('longitude') is None):raise ValueError('Renseigner les deux coordonnées ou aucune.')
        if p.get('latitude') is not None:p['latitude']=number(p['latitude'],'Latitude',-85,85);p['longitude']=number(p['longitude'],'Longitude',-180,180)
        validate_media(p.get('photo',''),700000,'Photo')
    quotes={}
    for q in data['quotes']:
        if q.get('project') not in project_ids:raise ValueError('Chantier du devis absent.')
        q['estimate']=estimate(q['estimate']['inputs']);day(q['date']);quotes[q['id']]=q
    totals={}
    for t in data['transactions']:
        q=quotes.get(t.get('quote'))
        if not q or t.get('project')!=q['project']:raise ValueError('Transaction sans devis correspondant.')
        if t.get('method') not in METHODS or t.get('status') not in STATUSES or t.get('mode')!='simulation':raise ValueError('Transaction de démonstration invalide.')
        t['amount']=number(t.get('amount'),'Montant',1,1e12,True)
        t['city']=text(t.get('city'),'Ville',80,True);text(t.get('date'),'Date',50,True)
        if t['status']=='Réussi':totals[q['id']]=totals.get(q['id'],0)+t['amount']
    for qid,total in totals.items():
        if total>quotes[qid]['estimate']['total']:raise ValueError('Le cumul des paiements simulés dépasse le devis.')
    for o in data['observations']:
        if o.get('project') not in project_ids:raise ValueError('Observation sans chantier.')
        if day(o['date'])>date.today():raise ValueError('Une observation ne peut pas être dans le futur.')
        o['note']=text(o.get('note'),'Observation',2000);o['progress']=number(o.get('progress'),'Avancement',0,100)
        validate_media(o.get('photo',''),700000,'Photo');validate_media(o.get('audio',''),8*1024*1024,'Note vocale')
        if o.get('audio_type') not in ('audio/wav','audio/mpeg','audio/mp4','audio/ogg','audio/webm'):raise ValueError('Format audio inconnu.')
    if len(json.dumps(data,ensure_ascii=False).encode())>MAX_BYTES:raise ValueError('Registre limité à 25 Mo. Réduire les pièces jointes et sauvegarder.')
    return data


def remaining(data,q):return q['estimate']['total']-sum(t['amount'] for t in data['transactions'] if t['quote']==q['id'] and t['status']=='Réussi')


def make_transaction(data,qid,method,amount):
    data=copy.deepcopy(data);q=next((q for q in data['quotes'] if q['id']==qid),None)
    if not q:raise ValueError('Sélectionner un devis enregistré.')
    amount=number(amount,'Montant du règlement',1,1e12,True)
    if amount>remaining(data,q):raise ValueError('Le montant dépasse le solde simulé restant.')
    if method not in METHODS:raise ValueError('Mode de paiement inconnu.')
    project=next(p for p in data['projects'] if p['id']==q['project'])
    data['transactions'].append(dict(id=new_code(data['transactions'],'TRX'),project=project['id'],quote=qid,
        city=project['city'],method=method,amount=amount,status='En attente',date=stamp(),mode='simulation'))
    return validate_store(data)


def settle(data,transaction_id,status):
    if status not in ('Réussi','Échoué','Annulé'):raise ValueError('Statut final invalide.')
    data=copy.deepcopy(data);t=next((t for t in data['transactions'] if t['id']==transaction_id),None)
    if not t:raise ValueError('Transaction absente.')
    if t['status']==status:return data
    if t['status']!='En attente':raise ValueError('Une transaction finalisée ne peut plus changer de statut.')
    t['status']=status;t['updated']=stamp()
    return validate_store(data)


def dataset(data):
    columns=['transaction','chantier','devis','ville','methode','montant_fcfa','statut','date','mode']
    return pd.DataFrame([[t[k] for k in ['id','project','quote','city','method','amount','status','date','mode']] for t in data['transactions']],columns=columns)


def csv_bytes(frame):
    frame=frame.copy()
    for col in frame.columns:
        if pd.api.types.is_string_dtype(frame[col]) or frame[col].dtype==object:
            frame[col]=frame[col].map(lambda x:"'"+x if isinstance(x,str) and x.lstrip().startswith(('=','+','-','@')) else x)
    return frame.to_csv(index=False,sep=';',decimal=',').encode('utf-8-sig')


def data_dir():
    default=Path.home()/('Documents/Codex/data/katika-btp-lab' if os.name=='nt' else '.local/share/katika-btp-lab')
    root=Path(os.environ.get('KATIKA_DATA_DIR',str(default))).expanduser().resolve();app=Path(__file__).resolve().parent
    if root==app or app in root.parents:raise ValueError('Le registre doit être conservé hors du dossier de l’application.')
    root.mkdir(parents=True,exist_ok=True);return root


def load_db():
    conn=sqlite3.connect(data_dir()/'katika.sqlite3',timeout=10)
    try:
        with conn:
            conn.execute('CREATE TABLE IF NOT EXISTS workspace (id INTEGER PRIMARY KEY CHECK(id=1),revision INTEGER,payload TEXT)')
            conn.execute('INSERT OR IGNORE INTO workspace VALUES(1,0,?)',(json.dumps(empty_store()),))
            rev,payload=conn.execute('SELECT revision,payload FROM workspace WHERE id=1').fetchone()
        return json.loads(payload),rev
    finally:conn.close()


def save_db(data,revision):
    data=validate_store(data);root=data_dir();conn=sqlite3.connect(root/'katika.sqlite3',timeout=10)
    try:
        with conn:
            cursor=conn.execute('UPDATE workspace SET payload=?,revision=revision+1 WHERE id=1 AND revision=?',(json.dumps(data,ensure_ascii=False,allow_nan=False),revision))
            if cursor.rowcount!=1:raise ValueError('Le registre a changé dans une autre session. Actualiser avant de réessayer.')
    finally:conn.close()
    # Projection reconstruite depuis le registre : aucun doublon sur changement de statut.
    warning=''
    try:
        target=root/'dataset_chantiers.csv';temporary=root/('dataset-'+uid()+'.tmp')
        temporary.write_bytes(csv_bytes(dataset(data)));temporary.replace(target)
    except OSError:warning='Données enregistrées. La copie CSV locale est indisponible ; utiliser l’export du dataset dans l’application.'
    return warning


def normalize_photo(raw):
    if len(raw)>5*1024*1024:raise ValueError('Photo limitée à 5 Mo.')
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('error',Image.DecompressionBombWarning)
            with Image.open(io.BytesIO(raw)) as source:
                if source.format not in ('JPEG','PNG','WEBP') or source.width*source.height>20000000:raise ValueError('Photo JPEG, PNG ou WebP, 20 millions de pixels maximum.')
                image=ImageOps.exif_transpose(source).convert('RGB');image.thumbnail((1200,900));out=io.BytesIO();image.save(out,'JPEG',quality=75,optimize=True)
    except Exception as exc:raise ValueError('Photo invalide ou trop grande.') from exc
    if out.tell()>700000:raise ValueError('Réduire la taille de cette photo.')
    return base64.b64encode(out.getvalue()).decode()


def normalize_audio(raw,filename):
    if len(raw)>8*1024*1024:raise ValueError('Note vocale limitée à 8 Mo.')
    ext=Path(filename).suffix.lower();types={'.wav':'audio/wav','.mp3':'audio/mpeg','.m4a':'audio/mp4','.ogg':'audio/ogg','.webm':'audio/webm'}
    if ext not in types:raise ValueError('Formats audio acceptés : WAV, MP3, M4A, OGG et WebM.')
    return base64.b64encode(raw).decode(),types[ext]


def demo_store():
    data=empty_store()
    p=new_project('Fondations · Villa Horizon','Yaoundé',600000);p.update(id='CE-2084',progress=35.,latitude=3.87,longitude=11.51)
    p2=new_project('Dallage · Atelier Delta','Douala',900000);p2.update(id='CE-2085',progress=15.)
    data['projects']=[p,p2]
    for index,project in enumerate(data['projects']):
        data['quotes'].append(dict(id=f'DEV-{index+1:04d}',project=project['id'],date=date.today().isoformat(),estimate=estimate(default_measure())))
    for i in range(12):
        project=p if i<10 else p2
        data['transactions'].append(dict(id=f'TRX-{i+1:04d}',project=project['id'],quote='DEV-0001' if i<10 else 'DEV-0002',
             city=project['city'],method='Orange Money' if i<7 else 'MTN MoMo',amount=10000,
             status='Réussi',date=date.today().isoformat(),mode='simulation'))
    data['observations']=[dict(id='OBS-0001',project=p['id'],date=date.today().isoformat(),created=stamp(),note='Exemple entièrement fictif : préparation des fondations.',progress=35.,photo='',audio='',audio_type='audio/wav')]
    return validate_store(data)


@st.cache_resource
def register_fonts():
    import reportlab
    root=Path(reportlab.__file__).resolve().parent/'fonts'
    for name,file in [('Katika','Vera.ttf'),('Katika-Bold','VeraBd.ttf')]:
        if name not in pdfmetrics.getRegisteredFontNames():pdfmetrics.registerFont(TTFont(name,str(root/file)))
    pdfmetrics.registerFontFamily('Katika',normal='Katika',bold='Katika-Bold',italic='Katika',boldItalic='Katika-Bold')


def quote_pdf(project,q,fictitious):
    register_fonts();styles=getSampleStyleSheet()
    for s in styles.byName.values():s.fontName='Katika-Bold' if s.name in ('Title','Heading2') else 'Katika'
    styles['Normal'].fontSize=9;styles['Normal'].leading=13
    P=lambda value:Paragraph(esc(value),styles['Normal'])
    e=q['estimate'];out=io.BytesIO();items=[Paragraph(APP_NAME,styles['Title']),P(f"DEVIS ESTIMATIF {q['id']} · {project['id']}"),P(project['name']+' · '+project['city']),P(q['date']+' · '+('EXEMPLE FICTIF' if fictitious else 'ESTIMATION UTILISATEUR')),Spacer(1,.5*cm)]
    rows=[[P(x) for x in ['Désignation','Quantité','Unité','Prix unitaire','Montant FCFA']]]
    rows += [[P(r['name']),P(fmt(r['quantity'],3)),P(r['unit']),P(fmt(r['price'],0)),P(fmt(r['amount']))] for r in e['rows']]
    rows += [[P('Main-d’œuvre / forfait'),P('1'),P('forfait'),P(fmt(e['inputs']['labor'])),P(fmt(e['inputs']['labor']))]]
    table=Table(rows,colWidths=[4.2*cm,2.2*cm,3.2*cm,3.4*cm,4.2*cm],repeatRows=1)
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#D9D1F5')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.whitesmoke,colors.white]),('VALIGN',(0,0),(-1,-1),'TOP'),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8)]))
    items += [table,Spacer(1,.4*cm),P('Sous-total : '+money(e['subtotal'])),P(f"Taxe saisie {e['inputs']['tax']:g} % : {money(e['tax'])}"),P('TOTAL : '+money(e['total'])),Spacer(1,.4*cm),P(f"Hypothèses : volume {fmt(e['volume'],3)} m³ ; dosage ciment {e['inputs']['dosage']:g} kg/m³ ; marge de quantités {e['inputs']['waste']:g} %. Les sacs de ciment sont arrondis au supérieur. Les autres unités ne sont pas additionnées."),P('Prix et coefficients saisis, sans validation réglementaire. Le dosage ne garantit pas une résistance du béton. Le ratio d’acier est une hypothèse de quantité, pas un calcul de ferraillage.')]
    def footer(c,d):c.setFont('Katika',8);c.drawString(1.4*cm,.8*cm,APP_NAME+' · Devis pédagogique');c.drawRightString(19.4*cm,.8*cm,str(d.page))
    SimpleDocTemplate(out,leftMargin=1.4*cm,rightMargin=1.4*cm,topMargin=1.5*cm,bottomMargin=1.5*cm,title=APP_NAME+' - Devis').build(items,onFirstPage=footer,onLaterPages=footer)
    return out.getvalue()


def excel_quotes(q):
    book=Workbook();s=book.active;s.title='Devis';s.append(['Désignation','Quantité','Unité','Prix unitaire FCFA','Montant FCFA'])
    for r in q['estimate']['rows']:s.append([r['name'],r['quantity'],r['unit'],r['price'],r['amount']])
    s.append(['Main-d’œuvre',1,'forfait',q['estimate']['inputs']['labor'],q['estimate']['inputs']['labor']]);s.append(['TOTAL',None,None,None,q['estimate']['total']])
    for column in s.columns:s.column_dimensions[column[0].column_letter].width=24
    out=io.BytesIO();book.save(out);return out.getvalue()


MATERIALS=[
 dict(id='sable',name='Sable',category='Granulats',definition='Granulat fin constitué de grains minéraux.',usage='Mortiers et bétons selon une formulation adaptée.',observation='Observer la propreté, la présence de fines et l’homogénéité. Ces observations ne remplacent pas les essais.',tags='granulat fin propreté granulométrie',color='#D1AD6B',source='https://www.usgs.gov/publications/natural-aggregates-conterminous-united-states'),
 dict(id='gravier',name='Gravier',category='Granulats',definition='Ensemble de grains grossiers naturels ou concassés.',usage='Bétons et couches granulaires selon les caractéristiques requises.',observation='Observer formes, propreté et diversité des grains ; les vérifier par des essais adaptés.',tags='granulat gros concassé',color='#9BA6B4',source='https://www.usgs.gov/publications/natural-aggregates-conterminous-united-states'),
 dict(id='calcaire',name='Calcaire',category='Roches',definition='Roche composée principalement de carbonate de calcium.',usage='Granulats, pierre de construction et matière première de liants.',observation='L’effervescence à un acide adapté peut indiquer des carbonates. Faire interpréter ce test par un encadrant ; il ne mesure pas la résistance.',tags='carbonate calcite test à l’acide',color='#C5C1AA',source='https://pubs.usgs.gov/fs/2008/3089/'),
 dict(id='gres',name='Grès',category='Roches',definition='Roche sédimentaire constituée de grains de sable liés.',usage='Pierre de construction ou parement selon ses propriétés.',observation='Observer les grains et la cohésion. Certains grès sont friables ; la roche doit être caractérisée.',tags='sédimentaire sable friable grès',color='#B68463',source='https://www.usgs.gov/faqs/what-are-sedimentary-rocks?qt-news_science_products=0'),
 dict(id='granite',name='Granite',category='Roches',definition='Roche magmatique présentant des cristaux imbriqués.',usage='Granulats et pierre de construction selon les caractéristiques mesurées.',observation='Observer les cristaux et l’altération ; l’aspect seul ne prouve pas une aptitude structurale.',tags='magmatique cristaux quartz',color='#9A94A0',source='https://floridadep.gov/fgs/geologic-topics/content/all-about-rocks'),
 dict(id='ciment',name='Ciment',category='Liants',definition='Liant hydraulique utilisé notamment dans les mortiers et bétons.',usage='Assemblage de constituants selon une formulation étudiée.',observation='Vérifier identification et état du conditionnement. Un dosage seul ne définit pas la résistance du béton.',tags='hydraulique liant dosage béton',color='#A6A69E',source='https://www.infociments.fr/sites/default/files/article/fichier/CT-T53.pdf'),
 dict(id='acier',name='Acier',category='Métalliques',definition='Matériau métallique employé notamment comme armature.',usage='Armatures ou éléments métalliques suivant un dimensionnement.',observation='Vérifier les références et l’état apparent ; les performances demandent documents et essais.',tags='fer armature métallique',color='#7D9BAD',source='https://www.infociments.fr/sites/default/files/article/fichier/CT-T53.pdf'),
]


@st.cache_data(max_entries=10)
def material_image(index):
    entry=MATERIALS[index];rng=np.random.default_rng(index+8);im=Image.new('RGB',(700,300),'#232838');draw=ImageDraw.Draw(im)
    for _ in range(240):
        x,y=rng.integers(0,700),rng.integers(0,260);size=int(rng.integers(3,24));draw.ellipse((int(x),int(y),int(x+size),int(y+size)),fill=entry['color'])
    draw.rectangle((0,260,700,300),fill='#101421');draw.text((18,275),'ILLUSTRATION PEDAGOGIQUE - PAS UNE PHOTO DE REFERENCE',fill='#DBDDED')
    out=io.BytesIO();im.save(out,'PNG');return out.getvalue()


GEO=st.components.v2.component('katika_geolocation',html='<button type="button">Utiliser ma position GPS</button><p role="status"></p>',css='button{background:#29253e;color:#e6dfff;border:1px solid #695987;border-radius:9px;padding:11px 15px;cursor:pointer;font:14px Arial}p{color:#c9bfdc;font:12px Arial}',js="""
export default function({parentElement,setStateValue}) {
 const button=parentElement.querySelector('button'),status=parentElement.querySelector('p');let disposed=false;
 const click=()=>{if(!navigator.geolocation){status.textContent='Géolocalisation indisponible.';return;}
 status.textContent='Demande de position…';navigator.geolocation.getCurrentPosition(p=>{if(disposed)return;status.textContent='Position reçue, précision '+Math.round(p.coords.accuracy)+' m.';setStateValue('position',{latitude:p.coords.latitude,longitude:p.coords.longitude,accuracy:p.coords.accuracy,id:Date.now()});},e=>{if(!disposed)status.textContent='Position non disponible. Tu peux saisir les coordonnées manuellement.';},{enableHighAccuracy:true,timeout:10000,maximumAge:0});};
 button.addEventListener('click',click);return ()=>{disposed=true;button.removeEventListener('click',click);};
}
""")


CSS='''<style>
.stApp,[data-testid="stAppViewContainer"]{background:#0D1019;color:#EFF1FA}
[data-testid="stHeader"]{background:#0D1019F0}.block-container{max-width:1470px;padding:4rem 2rem 3rem}
h1,h2,h3,h4{color:#F1F2FA!important;letter-spacing:-.025em}p,label{overflow-wrap:anywhere}
[data-testid="stMain"] [data-testid="stMarkdownContainer"]{color:#E6E8F4}
[data-testid="stMain"] [data-testid="stCaptionContainer"],[data-testid="stMain"] [data-testid="stCaptionContainer"] p{color:#ADB5CA!important}
[data-testid="stMain"] input,[data-testid="stMain"] textarea{color:#F1F1F9!important;background:#232838!important}
[data-testid="stMain"] [data-testid="stWidgetLabel"] p{color:#D9DEEC!important}
.stButton button,.stFormSubmitButton button,.stDownloadButton button,.stLinkButton a{border-radius:9px;min-height:42px}
button p{color:inherit!important}[data-testid="stMain"] button[kind="primary"]{background:#6F57C3;border-color:#8F78DF;color:#fff}
[data-testid="stMain"] button[kind="secondary"]{background:#1E2433;border:1px solid #353D54;color:#E5E8F3}
[data-testid="stForm"]{background:#171C2A;border-color:#333A50;border-radius:15px}
[data-testid="stExpander"]{border-color:#343B50;border-radius:12px}
[data-testid="stDataFrame"]{border:1px solid #323A50;border-radius:10px}
.kl-mast{display:flex;justify-content:space-between;align-items:center;gap:20px;padding:19px 23px;background:#171B28;border:1px solid #30364B;border-radius:14px;margin-bottom:15px}
.kl-wordmark{display:flex;align-items:center;gap:15px}.kl-mark{height:48px;width:48px;background:#9A83F5;clip-path:polygon(50% 0,94% 25%,94% 75%,50% 100%,6% 75%,6% 25%);position:relative}
.kl-mark:before{content:'';position:absolute;inset:5px;background:#181D2D;clip-path:inherit}.kl-mark:after{content:'K';position:absolute;inset:0;display:grid;place-items:center;color:#B2A0FF;font-weight:850;font-size:24px}
.kl-wordmark h1{font:800 clamp(19px,2.7vw,32px) Arial;margin:0;padding:0;color:#FAFAFF!important;letter-spacing:-.8px}
.kl-wordmark p{margin:4px 0 0;color:#ADB5CD;font-size:12px}.kl-workspace{background:#2A263B;border:1px solid #443A5E;border-radius:20px;padding:9px 15px;color:#CFC6EE;font-size:11px;max-width:240px}
.kl-dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:#3AD5A3;margin-right:6px}
.kl-tag{display:inline-block;color:#B7A7F0;background:#242036;padding:5px 9px;border-radius:5px;font-size:10px;font-weight:700;letter-spacing:.4px}
.kl-card-title{font:700 17px Arial;color:#F2F2FA;line-height:1.4;margin-bottom:10px}
.kl-big{font:850 clamp(36px,4vw,58px) Arial;color:#FF8B25;letter-spacing:-2px;line-height:1.35;margin:4px 0 8px;overflow-wrap:anywhere}
.kl-small{font-size:11px;color:#B0B8D0;line-height:1.6}.kl-statline{font-size:13px;color:#C7BDDE;margin:7px 0}.kl-statline strong{font-size:21px;color:#F3F3FF}
.st-key-overview_mc,.st-key-overview_corr,.st-key-overview_joint{background:#171B28;border:1px solid #343A50;border-radius:14px;padding:17px;min-height:405px}
.st-key-payment_box{background:linear-gradient(120deg,#1A2030,#191B27);border:1px solid #373E54;border-radius:15px;padding:18px;margin-top:18px}
.kl-pay-title{font-size:20px;font-weight:800;color:#F5F5FA;margin-bottom:6px}
.kl-pay-logo{display:inline-grid;place-items:center;height:42px;min-width:42px;border-radius:50%;background:var(--accent);color:#111826;font-size:12px;font-weight:900;margin-right:10px}
.kl-pay-amount{font-size:clamp(24px,3vw,35px);font-weight:850;color:var(--accent);letter-spacing:-1px;margin:8px 0;overflow-wrap:anywhere}
.st-key-orange_payment button{background:#FF8B25!important;color:#151722!important;border-color:#FF8B25!important;font-weight:800}
.st-key-momo_payment button{background:#FFD226!important;color:#151722!important;border-color:#FFD226!important;font-weight:800}
.kl-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:15px 0 20px}
.kl-kpi{border:1px solid #343C53;border-radius:12px;background:#191E2C;padding:17px;min-width:0}.kl-kpi small{color:#B6C0D5;font-size:11px}.kl-kpi strong{display:block;font-size:25px;color:#F1F3FC;margin:8px 0;overflow-wrap:anywhere}.kl-kpi span{font-size:10px;color:#A2ADC6}
.kl-note{border-left:3px solid #A089FF;background:#211D32;color:#CDC3E9;padding:13px 15px;font-size:12px;line-height:1.7;border-radius:0 9px 9px 0;margin:12px 0}
.kl-project{background:#1A2030;border:1px solid #363F56;border-radius:12px;padding:17px;min-height:110px}.kl-project strong{font-size:18px;color:#F5F5FD}.kl-project small{display:block;color:#B4BFD6;margin:8px 0;font-size:12px}
.kl-footer{color:#8E9BB9;text-align:center;font-size:10px;border-top:1px solid #293147;padding:18px;margin-top:25px}
@media(max-width:1000px){.kl-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.kl-mast{padding:17px}.kl-workspace{max-width:180px}.kl-wordmark h1{font-size:24px}}
@media(max-width:640px){.block-container{padding:4rem .85rem 2rem}.kl-mast{align-items:flex-start;flex-direction:column;gap:12px}.kl-mark{height:38px;width:38px;flex-shrink:0}.kl-wordmark{gap:10px}.kl-wordmark h1{font-size:22px}.kl-workspace{max-width:none}.kl-big{font-size:47px}.kl-kpi{padding:13px}.kl-kpi strong{font-size:21px}.st-key-overview_mc,.st-key-overview_corr,.st-key-overview_joint{min-height:0;padding:15px}.kl-pay-amount{font-size:28px}}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style>'''


def current():return st.session_state['kl_demo'] if st.session_state['kl_workspace']=='Exemple fictif' else st.session_state['kl_real']


def save(data,message='Données enregistrées.'):
    data=validate_store(data)
    if st.session_state['kl_workspace']=='Exemple fictif':st.session_state['kl_demo']=data
    else:
        warning=save_db(data,st.session_state['kl_revision'])
        if warning:message=warning
    st.session_state['kl_flash']=message


def commit(data,message='Données enregistrées.'):
    try:save(data,message);st.rerun()
    except (ValueError,OSError,sqlite3.Error) as exc:st.error(str(exc))


def load_example():
    st.session_state['kl_demo']=demo_store();st.session_state['kl_workspace']='Exemple fictif';st.session_state['_kl_workspace']='Exemple fictif';st.session_state['kl_project']='CE-2084';st.session_state['_kl_project']='CE-2084'
    st.session_state['kl_page']='Tableau de bord'


def change_workspace():
    st.session_state['kl_workspace']=st.session_state['_kl_workspace'];st.session_state.pop('kl_project',None)


def select_project():st.session_state['kl_project']=st.session_state['_kl_project']
def navigate(name):st.session_state['kl_page']=name


def selected_project(data):return next((p for p in data['projects'] if p['id']==st.session_state.get('kl_project')),None)


def require_project(data):
    p=selected_project(data)
    if not p:
        st.info('Crée un chantier dans Carnet de chantier, ou charge l’exemple fictif pour découvrir les outils.')
        if st.button('Créer mon premier chantier',type='primary'):navigate('Carnet de chantier');st.rerun()
    return p


def last_quote(data,p):return next((q for q in reversed(data['quotes']) if q['project']==p['id']),None)


def risk_context(data,p):
    q=last_quote(data,p);reference=q['estimate']['total'] if q else 500000.
    weights=[.25,.3,.15]
    if q and reference:
        amounts={r['key']:r['amount'] for r in q['estimate']['rows']}
        weights=[amounts[k]/reference for k in ['cement','steel','sand']]
    pref=posterior(data['transactions'],p['city'],p['risk'])
    result=simulate(p['budget'],reference,weights,p['risk'],pref['mean'])
    return result,pref,reference,weights


def kpis(items):
    st.html('<div class="kl-kpis">'+''.join(f'<div class="kl-kpi"><small>{esc(label)}</small><strong>{esc(value)}</strong><span>{esc(note)}</span></div>' for label,value,note in items)+'</div>')


def plot(fig,alt,height=300,legend=True):
    fig.update_layout(template='plotly_dark',paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',height=height,
        font=dict(family='Arial',color='#C7D0E5',size=11),margin=dict(l=15,r=15,t=24,b=20),showlegend=legend,
        legend=dict(orientation='h',y=-.18),colorway=[ORANGE,PURPLE,BLUE,YELLOW])
    fig.update_xaxes(gridcolor='#30374B',automargin=True,zerolinecolor='#59617A');fig.update_yaxes(gridcolor='#30374B',automargin=True,zerolinecolor='#59617A')
    st.plotly_chart(fig,width='stretch',theme=None,config={'displaylogo':False,'responsive':True},alt=alt)


def histogram_figure(result):
    h,edges=np.histogram(result['balance']/1000,bins=35)
    fig=go.Figure(go.Scatter(x=(edges[1:]+edges[:-1])/2,y=h,mode='lines',fill='tozeroy',line=dict(color=ORANGE,width=2),fillcolor='rgba(255,139,37,.35)',hovertemplate='%{x:.1f} milliers FCFA<br>%{y} scénarios<extra></extra>'))
    fig.add_vline(x=0,line_color='#D9C9FA',line_dash='dot');fig.update_layout(xaxis_title='R − S (milliers FCFA)',yaxis_title='Tirages')
    return fig


def heatmap(matrix,title='Corrélation hypothétique'):
    labels=['Ciment','Acier','Sable'];texts=[[fmt(v,2) if np.isfinite(v) else '—' for v in row] for row in matrix]
    fig=go.Figure(go.Heatmap(z=matrix,x=labels,y=labels,text=texts,texttemplate='%{text}',textfont=dict(size=14),
        colorscale=[[0,'#135D9A'],[.5,'#4F647F'],[1,'#E8953A']],zmin=-1 if np.nanmin(matrix)<0 else 0,zmax=1,showscale=False,hovertemplate='%{y} / %{x} : %{z:.3f}<extra></extra>'))
    fig.update_yaxes(autorange='reversed');return fig


def ellipse_figure(rho):
    cov=np.array([[1,rho],[rho,1]]);vals,vecs=np.linalg.eigh(cov);root=vecs@np.diag(np.sqrt(vals));radius=np.sqrt(stats.chi2.ppf(.95,2))
    theta=np.linspace(0,2*np.pi,150);outline=root@np.vstack([radius*np.cos(theta),radius*np.sin(theta)])
    fig=go.Figure(go.Scatter(x=outline[0],y=outline[1],mode='lines',fill='toself',fillcolor='rgba(160,137,255,.28)',line=dict(color=PURPLE,width=2),hoverinfo='skip'))
    for level in np.linspace(-radius*.85,radius*.85,9):
        width=np.sqrt(radius**2-level**2);v=np.linspace(-width,width,35)
        for points in [np.vstack([np.full_like(v,level),v]),np.vstack([v,np.full_like(v,level)])]:
            xy=root@points;fig.add_scatter(x=xy[0],y=xy[1],mode='lines',line=dict(color='rgba(185,167,255,.35)',width=1),hoverinfo='skip')
    fig.update_layout(xaxis_title='Score ressources Zᵣ',yaxis_title='Score coût Z꜀');fig.update_yaxes(scaleanchor='x',scaleratio=1)
    return fig


def payment_panel(data,p,q,dashboard=False):
    with st.container(key='payment_box'):
        st.html('<div class="kl-pay-title">Paiement · Devis et chantier</div>')
        if not q:
            st.caption('Enregistre un devis pour préparer une transaction de démonstration.');return
        balance=remaining(data,q)
        st.caption(f"Devis {q['id']} · Chantier {p['id']} · Émis le {q['date']} · {'Solde ouvert' if balance else 'Réglé en simulation'}")
        st.html('<div class="kl-note">MODE DÉMONSTRATION · aucune somme réelle n’est débitée. Les transactions alimentent le jeu de données des paiements simulés.</div>')
        amount_key='pay_amount_'+q['id'];balance_key='pay_balance_'+q['id']
        if st.session_state.get(balance_key)!=balance:
            st.session_state[amount_key]=int(balance);st.session_state[balance_key]=balance
        amount=st.number_input('Montant à régler en simulation (FCFA)',0,10**12,step=1000,key=amount_key,
                               help='Le montant est revérifié par rapport au solde du devis lors de la création.')
        for col,(method,accent,key) in zip(st.columns(2),[('Orange Money',ORANGE,'orange_payment'),('MTN MoMo',YELLOW,'momo_payment')]):
            with col:
                with st.container(key=key):
                    mark='OM' if method=='Orange Money' else 'MoMo'
                    st.html(f'<div><span class="kl-pay-logo" style="--accent:{accent}">{mark}</span><strong>{method}</strong><div class="kl-pay-amount" style="--accent:{accent}">{money(balance)}</div><div class="kl-small">Solde restant du devis · simulation</div></div>')
                    if st.button('Simuler un paiement '+method,disabled=balance<=0 or amount<=0,width='stretch',wrap=True):
                        try:
                            updated=make_transaction(data,q['id'],method,amount);save(updated,'Transaction créée au statut En attente.');navigate('Paiements');st.rerun()
                        except (ValueError,OSError,sqlite3.Error) as exc:st.error(str(exc))


def dashboard(data,p):
    if not p:
        for col,title,description in zip(st.columns(3),['Simulations Monte-Carlo','Corrélation des matériaux','Ressources et coûts'],['Probabilités calculées à partir de tes paramètres.','Ciment, acier et sable : dépendances explicites.','Une représentation des hypothèses du modèle.']):
            with col:
                with st.container(border=True):st.subheader(title);st.caption(description)
        require_project(data);return
    result,pref,reference,weights=risk_context(data,p);r=p['risk'];n=r['count']
    a,b,c=st.columns([1,1,1.05],gap='medium')
    with a:
        with st.container(key='overview_mc'):
            st.html(f'<div class="kl-card-title">Simulation Monte-Carlo</div><div class="kl-big">{fmt(100*result["probability"],1)} %</div><span class="kl-tag">PROBABILITÉ DE DÉFICIT : S &gt; R</span>')
            plot(histogram_figure(result),'Distribution simulée de la différence ressources moins dépenses',175,False)
            st.html(f'<div class="kl-small">{fmt(n)} scénarios · {result["hits"]} déficits observés</div><div class="kl-statline">Loi Bêta : <strong>α {fmt(r["alpha"],1)} · β {fmt(r["beta"],1)}</strong></div>')
    with b:
        with st.container(key='overview_corr'):
            st.html('<div class="kl-card-title">Matrice de corrélation</div><span class="kl-tag">CIMENT · ACIER · SABLE</span>')
            plot(heatmap(result['latent']),'Matrice hypothétique des corrélations des scores de prix',250,False)
            st.html(f'<div class="kl-statline">Ciment–acier : <strong>{fmt(r["rho_cs"],2)}</strong></div><div class="kl-small">Hypothèses modifiables · données simulées</div>')
    with c:
        with st.container(key='overview_joint'):
            st.html('<div class="kl-card-title">Dépendance ressources / coûts</div><span class="kl-tag">SCORES GAUSSIENS DU MODÈLE</span>')
            plot(ellipse_figure(r['rho_rc']),'Ellipse à 95 pour cent des deux scores gaussiens latents',240,False)
            st.html(f'<div class="kl-statline">Ellipse à 95 % · <strong>ρ {fmt(r["rho_rc"],2)}</strong></div><div class="kl-small">R et S sont des montants transformés. L’ellipse concerne leurs scores, pas directement les FCFA.</div>')
    payment_panel(data,p,last_quote(data,p),True)
    kpis([('Chantier actif',p['id'],p['city']),('Avancement déclaré',fmt(p['progress'],1)+' %','dernier état du carnet'),('Devis de référence',money(reference),'dernier devis ou hypothèse 500 000'),('Préférence OM · ville',fmt(pref['mean']*100,1)+' %',f"{pref['count']} règlements simulés réussis · moyenne a posteriori")])
    if st.button('Examiner les paramètres et les données',icon=':material/query_stats:'):navigate('Risques et données');st.rerun()


def measure_page(data,p):
    st.header('Métrés et devis');st.caption('Calculs géométriques et estimation en FCFA. Les coefficients et prix sont des hypothèses modifiables.')
    if not p:require_project(data);return
    previous=last_quote(data,p);x=copy.deepcopy(previous['estimate']['inputs'] if previous else default_measure())
    with st.form('measure_'+p['id']):
        x['label']=st.text_input('Désignation des travaux',x['label'],max_chars=120)
        a,b,c=st.columns(3)
        for col,key,label in [(a,'length','Longueur (m)'),(b,'width','Largeur (m)'),(c,'height','Hauteur / épaisseur (m)')]:
            with col:x[key]=st.number_input(label,.001,1000. if key=='height' else 10000.,float(x[key]),format='%.3f')
        a,b,c=st.columns(3)
        with a:x['dosage']=st.number_input('Dosage ciment (kg/m³)',0.,1000.,float(x['dosage']),step=50.);x['bag']=st.number_input('Masse d’un sac (kg)',1.,100.,float(x['bag']))
        with b:x['sand']=st.number_input('Sable (m³ par m³ de béton)',0.,3.,float(x['sand']));x['gravel']=st.number_input('Gravier (m³ par m³ de béton)',0.,3.,float(x['gravel']))
        with c:x['water_ratio']=st.number_input('Rapport eau / ciment (L par kg)',0.,2.,float(x['water_ratio']));x['steel']=st.number_input('Hypothèse d’acier (kg/m³)',0.,1000.,float(x['steel']))
        x['waste']=st.number_input('Marge sur les quantités (%)',0.,100.,float(x['waste']))
        st.caption('150, 250 et 350 kg/m³ sont des dosages possibles à explorer. Un dosage seul ne garantit pas une résistance ; l’acier n’est pas dimensionné par cet outil.')
        with st.expander('Prix unitaires et frais',expanded=True):
            a,b=st.columns(2)
            for i,(key,label) in enumerate([('price_cement','Prix d’un sac de ciment (FCFA)'),('price_sand','Prix du sable (FCFA/m³)'),('price_gravel','Prix du gravier (FCFA/m³)'),('price_water','Prix de l’eau (FCFA/L)'),('price_steel','Prix de l’acier (FCFA/kg)'),('labor','Main-d’œuvre et forfaits (FCFA)')]):
                with (a if i%2==0 else b):x[key]=st.number_input(label,0.,1e10 if key=='labor' else 1e8,float(x[key]))
            x['tax']=st.number_input('Taxe saisie, si applicable (%)',0.,100.,float(x['tax']))
        submitted=st.form_submit_button('Calculer et enregistrer le devis',type='primary',width='stretch')
    if submitted:
        try:
            e=estimate(x);updated=copy.deepcopy(data);q=dict(id=new_code(data['quotes'],'DEV'),project=p['id'],date=date.today().isoformat(),estimate=e);updated['quotes'].append(q);commit(updated,'Devis enregistré. Les analyses utilisent désormais ce devis de référence.')
        except ValueError as exc:st.error(str(exc))
    q=last_quote(data,p)
    if q:
        e=q['estimate'];kpis([('Surface',fmt(e['surface'],2)+' m²','longueur × largeur'),('Volume',fmt(e['volume'],3)+' m³','surface × épaisseur'),('Ciment théorique',fmt(e['theoretical']['cement'],1)+' kg','avant marge et arrondi des sacs'),('Total estimé',money(e['total']),q['id'])])
        st.dataframe(pd.DataFrame([{'Matériau':r['name'],'Quantité':r['quantity'],'Unité':r['unit'],'Prix unitaire (FCFA)':r['price'],'Montant (FCFA)':r['amount']} for r in e['rows']]),hide_index=True,width='stretch',alt='Quantités et coûts calculés du dernier devis enregistré')
        st.caption(f"Achats : {e['bags']} sacs de {e['inputs']['bag']:g} kg. Le ciment théorique est distinct de la quantité achetée après marge et arrondi au sac entier.")
        a,b=st.columns(2)
        with a:st.download_button('Télécharger le devis PDF',quote_pdf(p,q,st.session_state['kl_workspace']=='Exemple fictif'),q['id']+'.pdf','application/pdf',width='stretch')
        with b:st.download_button('Télécharger le devis Excel',excel_quotes(q),q['id']+'.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',width='stretch')
        message=f"Devis {q['id']} - {p['name']} : {money(e['total'])}. Le PDF est à joindre séparément."
        st.link_button('Préparer un message WhatsApp','https://wa.me/?text='+urlquote(message),icon=':material/share:')
        st.caption('Ce lien ouvre un message à compléter. Aucun message ni fichier n’est envoyé automatiquement.')
        with st.expander('Historique des devis'):
            st.dataframe(pd.DataFrame([{'Référence':z['id'],'Date':z['date'],'Total (FCFA)':z['estimate']['total']} for z in data['quotes'] if z['project']==p['id']]),hide_index=True,width='stretch',alt='Historique des devis du chantier')


def risk_page(data,p):
    st.header('Risques, vecteurs aléatoires et données')
    if not p:require_project(data);return
    r=copy.deepcopy(p['risk'])
    with st.form('risk_'+p['id']):
        st.markdown('**Ressources et dépenses**')
        a,b,c=st.columns(3)
        with a:budget=st.number_input('Ressources moyennes disponibles (FCFA)',0.,1e12,float(p['budget']));r['resource_cv']=st.number_input('Variation des ressources : CV (%)',0.,100.,r['resource_cv']*100)/100
        with b:r['cost_cv']=st.number_input('Variation globale du coût : CV (%)',0.,100.,r['cost_cv']*100)/100;r['rho_rc']=st.number_input('Corrélation latente ressources / coût',-.99,.99,float(r['rho_rc']))
        with c:
            choices=sorted(set([1000,10000,100000,r['count']]))
            r['count']=st.selectbox('Nombre de tirages',choices,index=choices.index(r['count']));r['seed']=st.number_input('Graine aléatoire',0,2_000_000_000,r['seed'])
        st.markdown('**Prix des matériaux et dépendances**')
        a,b,c=st.columns(3)
        for col,key,label,rho,labelrho in [(a,'cv_cement','CV ciment (%)','rho_cs','Corrélation ciment / acier'),(b,'cv_steel','CV acier (%)','rho_csa','Corrélation ciment / sable'),(c,'cv_sand','CV sable (%)','rho_ssa','Corrélation acier / sable')]:
            with col:r[key]=st.number_input(label,0.,100.,r[key]*100)/100;r[rho]=st.number_input(labelrho,-1.,1.,float(r[rho]))
        st.markdown('**Surcoût incident : occurrence × intensité Bêta**')
        a,b=st.columns(2)
        with a:r['incident_probability']=st.number_input('Probabilité d’incident (%)',0.,100.,r['incident_probability']*100)/100;r['alpha']=st.number_input('Paramètre α de la loi Bêta',.1,100.,float(r['alpha']))
        with b:r['max_extra']=st.number_input('Surcoût maximum de l’incident (FCFA)',0.,1e10,float(r['max_extra']));r['beta']=st.number_input('Paramètre β de la loi Bêta',.1,100.,float(r['beta']))
        with st.expander('A priori et frais hypothétiques de paiement'):
            a,b=st.columns(2)
            with a:r['prior_om']=st.number_input('A priori Orange Money (α)',.1,1000.,float(r['prior_om']));r['fee_om']=st.number_input('Frais Orange Money supposés (%)',0.,10.,float(r['fee_om']))
            with b:r['prior_momo']=st.number_input('A priori MTN MoMo (β)',.1,1000.,float(r['prior_momo']));r['fee_momo']=st.number_input('Frais MTN MoMo supposés (%)',0.,10.,float(r['fee_momo']))
            st.caption('Frais nuls par défaut : aucun tarif opérateur n’est supposé. Les frais saisis ne concernent que ce scénario et ne sont pas facturés.')
        run=st.form_submit_button('Enregistrer et lancer la simulation',type='primary',width='stretch')
    if run:
        try:
            r=validate_risk(r);updated=copy.deepcopy(data);target=next(x for x in updated['projects'] if x['id']==p['id']);target.update(risk=r,budget=budget);commit(updated,'Hypothèses enregistrées. Simulation recalculée.')
        except ValueError as exc:st.error(str(exc))
    result,pref,reference,weights=risk_context(data,p);lo,hi=result['interval']
    kpis([('P(S > R)',fmt(result['probability']*100,2)+' %',f"{result['hits']} déficits / {p['risk']['count']} tirages"),('Marge médiane R − S',money(np.median(result['balance'])),'montants simulés'),('Coût de référence',money(reference),'dernier devis, sinon hypothèse'),('Indice β',fmt(result['beta_index'],2) if result['beta_index'] is not None else 'Non estimé','−Φ⁻¹(Pdéficit), distinct de la loi Bêta')])
    st.caption(f"Intervalle de Wilson à 95 % pour la fréquence Monte-Carlo : {fmt(lo*100,2)} à {fmt(hi*100,2)} %. Il ne mesure pas la validité des hypothèses. Il s’agit d’un déficit à l’horizon, pas d’une probabilité de ruine d’un processus suivi dans le temps.")
    plot(histogram_figure(result),'Histogramme des marges simulées ressources moins dépenses',320,False)
    a,b=st.columns(2)
    with a:
        st.subheader('Loi Bêta de l’intensité d’incident')
        xs=np.linspace(.001,.999,220);fig=go.Figure(go.Scatter(x=xs,y=stats.beta.pdf(xs,p['risk']['alpha'],p['risk']['beta']),fill='tozeroy',line=dict(color=PURPLE),name='Densité Bêta'))
        fig.update_layout(xaxis_title='Intensité entre 0 et 1',yaxis_title='Densité');plot(fig,'Densité de la loi Bêta choisie',300,False)
    with b:
        st.subheader('Covariance des variations de prix')
        labels=['Ciment','Acier','Sable'];mat=result['covariance']
        fig=go.Figure(go.Heatmap(z=mat,x=labels,y=labels,text=[[fmt(v,2) for v in row] for row in mat],texttemplate='%{text}',colorscale='Blues',showscale=False));fig.update_yaxes(autorange='reversed')
        plot(fig,'Covariance simulée des variations de prix en points de pourcentage au carré',300,False)
    st.subheader('Préférences de paiement par ville : une mise à jour bayésienne')
    if pref['count']:
        st.markdown(f"À **{p['city']}**, {pref['om']} règlements Orange Money et {pref['momo']} MTN MoMo sont réussis en simulation : fréquence observée **{fmt(pref['empirical']*100,1)} %** pour OM. Avec l’a priori choisi, la moyenne a posteriori vaut **{fmt(pref['mean']*100,1)} %**.")
    else:st.info('Aucun règlement réussi dans cette ville. L’estimation utilise uniquement l’a priori, sans observation locale.')
    st.caption(f"Posterior : Bêta({fmt(pref['alpha'],1)}, {fmt(pref['beta'],1)}), intervalle de crédibilité à 95 % : {fmt(pref['interval'][0]*100,1)} à {fmt(pref['interval'][1]*100,1)} %. Ces données de paiement sont simulées ; elles ne décrivent pas la population d’une ville.")
    simulated=pd.DataFrame({'Ressources R (FCFA)':result['resource'],'Dépenses S (FCFA)':result['expenses'],'Marge R-S (FCFA)':result['balance'],'Moyen M simulé':np.where(result['method_om'],'Orange Money','MTN MoMo')})
    st.download_button('Exporter les tirages CSV',csv_bytes(simulated),'katika_simulations.csv','text/csv')
    with st.expander('Comprendre les formules et les limites'):
        st.markdown('R suit une loi lognormale de moyenne égale aux ressources saisies. Le coût de base est le dernier devis : les postes ciment, acier et sable subissent des multiplicateurs lognormaux corrélés, puis un facteur global de coût est appliqué. S ajoute un incident éventuel, dont l’intensité suit Bêta(α,β), et les frais hypothétiques liés à M. Leurs paramètres sont des hypothèses, pas un ajustement sur des données de marché.')
        st.markdown('M est une catégorie, pas une troisième variable normale. Les scores gaussiens servent à construire la dépendance ; leur corrélation est différente de celle des montants transformés. La covariance affichée concerne les variations de prix simulées. Une corrélation n’établit pas une causalité.')


def encyclopedia_page(data,p):
    import re,unicodedata
    st.header('Encyclopédie BTP');st.caption('Fiches pédagogiques embarquées, accessibles sans Internet sur un serveur local. Les liens de sources nécessitent une connexion.')
    a,b=st.columns([1.4,1])
    with a:query=st.text_input('Rechercher un matériau ou un indice',placeholder='Exemples : calcaire, test à l’acide, fer, granulats')
    with b:category=st.selectbox('Catégorie',['Toutes','Granulats','Roches','Liants','Métalliques'])
    def norm(s):return ''.join(c for c in unicodedata.normalize('NFKD',s.casefold()) if not unicodedata.combining(c))
    words=[w for w in re.findall(r'\w+',norm(query)) if w not in ['a','l','de','du','le','la','les']]
    rows=[(i,m) for i,m in enumerate(MATERIALS) if (category=='Toutes' or m['category']==category) and all(w in norm(' '.join(str(v) for v in m.values())) for w in words)]
    if not rows:st.info('Aucune fiche ne correspond à cette recherche.')
    for start in range(0,len(rows),3):
        for col,(index,entry) in zip(st.columns(3),rows[start:start+3]):
            with col:
                with st.container(border=True):
                    st.image(material_image(index),width='stretch',alt='Illustration pédagogique de '+entry['name'])
                    st.subheader(entry['name']);st.caption(entry['category']);st.markdown(entry['definition'])
                    with st.expander('Usage et observation terrain'):
                        st.markdown('**Usage :** '+entry['usage']);st.markdown('**Observation :** '+entry['observation']);st.link_button('Consulter la source',entry['source'])
    with st.expander('Fiches texte à conserver hors connexion'):
        content='KATIKA DATA LAB - FICHES PEDAGOGIQUES\n\n'+'\n\n'.join(m['name']+'\n'+m['definition']+'\nUsage : '+m['usage']+'\nObservation : '+m['observation']+'\nSource : '+m['source'] for m in MATERIALS)
        st.download_button('Télécharger les fiches texte',content.encode('utf-8'),'katika_fiches_materiaux.txt','text/plain')
        st.caption('Le fichier téléchargé est lisible hors connexion. Le site nécessite un serveur accessible ; il ne s’agit pas d’une application mobile autonome hors ligne.')


def payments_page(data,p):
    st.header('Paiements et collecte de données');st.caption('Prototype de paiement simulé. Les boutons ne déclenchent aucune transaction Orange Money ou MTN réelle.')
    if not p:require_project(data);return
    quotes=[q for q in data['quotes'] if q['project']==p['id']]
    if quotes:
        qid=st.selectbox('Devis à régler',[q['id'] for q in quotes],index=len(quotes)-1);q=next(q for q in quotes if q['id']==qid)
        payment_panel(data,p,q)
    else:st.info('Enregistrer d’abord un devis dans Métrés et devis.')
    relevant=[t for t in data['transactions'] if t['project']==p['id']]
    if relevant:
        st.subheader('Suivre une transaction')
        tid=st.selectbox('Référence transaction',[t['id'] for t in reversed(relevant)])
        t=next(t for t in relevant if t['id']==tid)
        st.markdown(f"**{t['id']}** · {t['method']} · **{money(t['amount'])}** · statut **{t['status']}**")
        if t['status']=='En attente':
            a,b,c=st.columns(3)
            for col,label,status in [(a,'Simuler la réussite','Réussi'),(b,'Simuler un échec','Échoué'),(c,'Annuler la simulation','Annulé')]:
                with col:
                    if st.button(label,width='stretch',wrap=True):
                        try:commit(settle(data,t['id'],status),'Statut mis à jour ; le dataset conserve une seule ligne par transaction.')
                        except ValueError as exc:st.error(str(exc))
        else:st.caption('Cette transaction est finalisée. Une nouvelle tentative crée une nouvelle référence.')
        frame=dataset(data);st.dataframe(frame[frame['chantier']==p['id']],hide_index=True,width='stretch',alt='Historique des transactions simulées du chantier')
    successful=[t for t in data['transactions'] if t['status']=='Réussi']
    if successful:
        counts=pd.DataFrame(successful).groupby(['city','method']).size().unstack(fill_value=0)
        fig=go.Figure()
        for method,col in zip(METHODS,[ORANGE,YELLOW]):
            fig.add_bar(x=counts.index,y=counts[method] if method in counts else [0]*len(counts),name=method,marker_color=col)
        fig.update_layout(barmode='group',xaxis_title='Ville des chantiers',yaxis_title='Nombre de règlements simulés réussis');plot(fig,'Répartition des moyens de paiement simulés par ville')
    st.download_button('Exporter dataset_chantiers.csv',csv_bytes(dataset(data)),'dataset_chantiers.csv','text/csv')
    st.caption('Le dataset contient les tentatives et leur statut courant. Seuls les règlements réussis alimentent les sommes réglées et les préférences. Les échecs sont conservés pour l’analyse des tentatives.')


def project_form(data,p=None):
    key=p['id'] if p else 'new'
    geo=GEO(key='gps_'+key,data={},width='stretch',height=72,on_position_change=lambda:None)
    position=getattr(geo,'position',None)
    if position and isinstance(position,dict) and position.get('id')!=st.session_state.get('gps_used_'+key):
        st.session_state['lat_'+key]=str(position['latitude']);st.session_state['lon_'+key]=str(position['longitude']);st.session_state['gps_used_'+key]=position['id']
        st.caption('Coordonnées proposées par le navigateur, à vérifier avant enregistrement.')
    with st.form('project_'+key):
        a,b=st.columns(2)
        with a:name=st.text_input('Nom du chantier',p['name'] if p else '',max_chars=100);city=st.text_input('Ville',p['city'] if p else 'Yaoundé',max_chars=80)
        with b:
            budget=st.number_input('Ressources prévues (FCFA)',0.,1e12,float(p['budget']) if p else 600000.)
            photo=st.file_uploader('Photo de présentation du chantier',type=['jpg','jpeg','png','webp'],max_upload_size=5)
        a,b=st.columns(2)
        with a:lat=st.text_input('Latitude (°), facultative',str(p['latitude']) if p and p['latitude'] is not None else '',key='lat_'+key)
        with b:lon=st.text_input('Longitude (°), facultative',str(p['longitude']) if p and p['longitude'] is not None else '',key='lon_'+key)
        submitted=st.form_submit_button('Enregistrer le chantier' if p else 'Créer le chantier',type='primary',width='stretch')
    if submitted:
        try:
            updated=copy.deepcopy(data);target=copy.deepcopy(p) if p else new_project(name,city,budget)
            target.update(name=text(name,'Nom du chantier',100,True),city=text(city,'Ville',80,True),budget=budget,
                          latitude=number(lat,'Latitude',-85,85) if lat.strip() else None,
                          longitude=number(lon,'Longitude',-180,180) if lon.strip() else None)
            if photo:target['photo']=normalize_photo(photo.getvalue())
            if not p:target['id']=new_code(data['projects'],'CE');updated['projects'].append(target)
            else:updated['projects']=[target if x['id']==p['id'] else x for x in updated['projects']]
            validate_store(updated);save(updated,'Chantier enregistré.');st.session_state['kl_project']=target['id'];st.rerun()
        except (ValueError,OSError,sqlite3.Error) as exc:st.error(str(exc))


def notebook_page(data,p):
    st.header('Carnet de chantier')
    for start in range(0,len(data['projects']),3):
        for col,project in zip(st.columns(3),data['projects'][start:start+3]):
            with col:
                st.html(f'<div class="kl-project"><span class="kl-tag">{esc(project["id"])}</span><br><strong>{esc(project["name"])}</strong><small>{esc(project["city"])} · {fmt(project["progress"],1)} %</small></div>')
                if st.button('Ouvrir '+project['id'],key='open_'+project['id'],width='stretch'):
                    st.session_state['kl_project']=project['id'];st.rerun()
    with st.expander('Créer un nouveau chantier',expanded=not data['projects']):project_form(data)
    if not p:return
    st.subheader(p['name']);st.progress(p['progress']/100,text='Avancement déclaré : '+fmt(p['progress'],1)+' %')
    if p['photo']:st.image(base64.b64decode(p['photo']),width='stretch',alt='Photo de présentation du chantier')
    if p['latitude'] is not None:
        st.map(pd.DataFrame([{'lat':p['latitude'],'lon':p['longitude']}]),zoom=12,alt='Localisation déclarée du chantier')
        st.caption(f"Coordonnées déclarées : {p['latitude']:.6f}, {p['longitude']:.6f}. Le fond de carte demande Internet.")
    with st.expander('Modifier les informations du chantier'):project_form(data,p)
    st.subheader('Ajouter une observation')
    with st.form('observation_'+p['id']):
        a,b=st.columns(2)
        with a:when=st.date_input('Date de l’observation',date.today(),max_value=date.today());progress=st.number_input('Avancement observé (%)',0.,100.,float(p['progress']))
        with b:photo=st.file_uploader('Photo de cette observation',type=['jpg','jpeg','png','webp'],max_upload_size=5)
        note=st.text_area('Observation écrite',max_chars=2000,height=90)
        recorded=st.audio_input('Enregistrer une note vocale',key='voice_'+p['id'])
        uploaded=st.file_uploader('Ou joindre une note vocale (8 Mo maximum)',type=['wav','mp3','m4a','ogg','webm'],max_upload_size=8)
        submitted=st.form_submit_button('Ajouter au carnet',type='primary',width='stretch')
    if submitted:
        try:
            audio=recorded or uploaded;blob,mime=normalize_audio(audio.getvalue(),audio.name) if audio else ('','audio/wav')
            if not note.strip() and not photo and not audio:raise ValueError('Ajouter une observation écrite, une photo ou une note vocale.')
            updated=copy.deepcopy(data);entry=dict(id=new_code(data['observations'],'OBS'),project=p['id'],date=when.isoformat(),created=stamp(),note=note,progress=progress,
                photo=normalize_photo(photo.getvalue()) if photo else '',audio=blob,audio_type=mime)
            updated['observations'].append(entry)
            latest=max((o for o in updated['observations'] if o['project']==p['id']),key=lambda o:(o['date'],o['created']))
            next(x for x in updated['projects'] if x['id']==p['id'])['progress']=latest['progress'];commit(updated,'Observation ajoutée au carnet.')
        except ValueError as exc:st.error(str(exc))
    observations=sorted([o for o in data['observations'] if o['project']==p['id']],key=lambda o:(o['date'],o['created']),reverse=True)
    if observations:
        fig=go.Figure(go.Scatter(x=[o['date'] for o in reversed(observations)],y=[o['progress'] for o in reversed(observations)],mode='lines+markers',line=dict(color=PURPLE)))
        fig.update_layout(xaxis_title='Date d’observation',yaxis_title='Avancement déclaré (%)',yaxis_range=[0,100]);plot(fig,'Historique des observations d’avancement')
        for o in observations[:20]:
            with st.expander(o['date']+' · '+o['id']+' · '+fmt(o['progress'],1)+' %',expanded=o==observations[0]):
                st.text(o['note'])
                if o['photo']:st.image(base64.b64decode(o['photo']),width='stretch',alt='Photo de l’observation '+o['id'])
                if o['audio']:st.audio(base64.b64decode(o['audio']),format=o['audio_type'])
        st.caption('Les 20 observations les plus récentes sont affichées ; la sauvegarde contient le carnet complet. Une note vocale est conservée et réécoutable, sans transcription automatique.')


def merge_backup(data,payload):
    if len(payload)>MAX_BYTES:raise ValueError('Sauvegarde limitée à 25 Mo.')
    try:backup=json.loads(payload)
    except Exception:raise ValueError('Fichier JSON illisible.') from None
    if not isinstance(backup,dict) or backup.get('application') not in (APP_NAME,'KATIKA BTP LAB'):raise ValueError('Sauvegarde KATIKA non reconnue.')
    incoming=validate_store(backup.get('data'));out=copy.deepcopy(data);projects={};quotes={}
    for p in incoming['projects']:
        old=p['id'];p['id']=new_code(out['projects'],'CE');projects[old]=p['id'];p['name']=text(p['name']+' (copie)','Nom',100,True);out['projects'].append(p)
    for q in incoming['quotes']:
        old=q['id'];q['id']=new_code(out['quotes'],'DEV');quotes[old]=q['id'];q['project']=projects[q['project']];out['quotes'].append(q)
    for t in incoming['transactions']:
        t['id']=new_code(out['transactions'],'TRX');t['project']=projects[t['project']];t['quote']=quotes[t['quote']];out['transactions'].append(t)
    for o in incoming['observations']:
        o['id']=new_code(out['observations'],'OBS');o['project']=projects[o['project']];out['observations'].append(o)
    return validate_store(out)


def data_page(data,p):
    st.header('Données, sauvegardes et guide')
    st.download_button('Télécharger dataset_chantiers.csv',csv_bytes(dataset(data)),'dataset_chantiers.csv','text/csv')
    backup=dict(application=APP_NAME,fictitious=st.session_state['kl_workspace']=='Exemple fictif',exported=stamp(),data=data)
    st.download_button('Sauvegarde complète JSON',json.dumps(backup,ensure_ascii=False,indent=2).encode('utf-8'),'katika_sauvegarde.json','application/json')
    with st.expander('Restaurer des copies depuis une sauvegarde'):
        with st.form('restore'):
            uploaded=st.file_uploader('Sauvegarde KATIKA JSON',type=['json'],max_upload_size=25)
            confirmed=st.checkbox('Ajouter des copies avec de nouvelles références')
            submitted=st.form_submit_button('Restaurer les copies')
        if submitted:
            try:
                if not uploaded or not confirmed:raise ValueError('Choisir une sauvegarde et confirmer l’ajout.')
                metadata=json.loads(uploaded.getvalue())
                if metadata.get('fictitious') and st.session_state['kl_workspace']=='Mes projets':raise ValueError('Restaurer cette sauvegarde fictive dans Exemple fictif.')
                commit(merge_backup(data,uploaded.getvalue()),'Copies restaurées sans supprimer les projets existants.')
            except (ValueError,AttributeError) as exc:st.error(str(exc))
    st.subheader('Parcours de test')
    cases=[['Accueil','Charger l’exemple fictif','Deux chantiers, trois cartes analytiques et deux méthodes de paiement.'],
           ['Métré','5 × 4 × 0,15 m, dosage 350, marge 0 %','Volume 3 m³ ; ciment théorique 1 050 kg ; 21 sacs de 50 kg.'],
           ['Devis','Garder tous les réglages initiaux','Marge 5 %, 23 sacs achetés ; total 498 581 FCFA.'],
           ['Données','Yaoundé dans l’exemple','7 OM / 3 MTN : fréquence OM 70 % ; posterior Bêta(8,4), moyenne 66,7 %.'],
           ['Paiements','Créer une transaction puis simuler la réussite','Une seule référence ; solde du devis diminué une seule fois.'],
           ['Carnet','Ajouter une note avec avancement 50 %','Observation conservée ; carte du chantier actualisée si c’est la plus récente.'],
           ['Risque exact','CV ressources/coût/prix = 0, incident = 0, frais = 0','Ressources 600 000 > devis 498 581 : aucun déficit simulé.'],
           ['Sauvegarde','Exporter puis restaurer dans l’exemple','Copies ajoutées avec nouveaux identifiants et relations conservées.']]
    st.dataframe(pd.DataFrame(cases,columns=['Outil','Saisie / action','Résultat attendu']),hide_index=True,width='stretch',alt='Parcours de test et résultats attendus')
    st.markdown('**Hors connexion :** le serveur local peut afficher les fiches embarquées sans Internet ; les cartes, liens de sources et WhatsApp en demandent. Sur téléphone, le site nécessite l’accès au serveur. Les fichiers téléchargés restent consultables hors ligne.')
    st.markdown('**Paiements réels :** non activés dans cette version. Un futur branchement exige des accès marchands, un serveur de confirmation et une vérification indépendante des statuts. Aucun frais ni débit réel n’est effectué ici.')
    st.markdown('**Stockage :** Mes projets est sauvegardé dans SQLite hors du dossier livré ; un fichier dataset_chantiers.csv y est reconstruit après chaque enregistrement. L’exemple fictif reste dans la session. Sur un hébergement au disque temporaire, conserver la sauvegarde JSON.')
    st.markdown('**Confidentialité des tests :** les photos et notes vocales sont des pièces jointes du registre. Les simulations emploient des paramètres hypothétiques et les paiements sont marqués simulation. Le dossier livré conserve app.py, requirement.text et README.')


def main():
    st.set_page_config(page_title=APP_NAME+' · Analyse de chantier',page_icon=':material/hub:',layout='wide')
    st.html(CSS)
    st.session_state.setdefault('kl_workspace','Mes projets');st.session_state.setdefault('kl_demo',empty_store());st.session_state.setdefault('kl_page','Tableau de bord')
    try:st.session_state['kl_real'],st.session_state['kl_revision']=load_db()
    except (ValueError,OSError,sqlite3.Error) as exc:st.error('Registre indisponible : '+str(exc));st.stop()
    data=current()
    if data['projects'] and not selected_project(data):st.session_state['kl_project']=data['projects'][0]['id']
    p=selected_project(data)
    st.html('<header class="kl-mast"><div class="kl-wordmark"><div class="kl-mark" aria-hidden="true"></div><div><h1>KATIKA DATA LAB</h1><p>Analyse · Simulation · Paiements · Carnet de chantier</p></div></div><div class="kl-workspace"><span class="kl-dot"></span>'+esc(st.session_state['kl_workspace'])+' · interface française</div></header>')
    with st.container(horizontal=True):
        for name in NAV:st.button(name,key='nav_'+name,type='primary' if name==st.session_state['kl_page'] else 'secondary',on_click=navigate,args=(name,),wrap=True)
    a,b,c=st.columns([1.45,1,.9],vertical_alignment='bottom')
    with a:
        if p:
            st.session_state['_kl_project']=p['id']
            lookup={x['id']:x['name'] for x in data['projects']}
            st.selectbox('Chantier actif',list(lookup),format_func=lambda x:x+' · '+lookup[x],key='_kl_project',on_change=select_project)
        else:st.caption('Aucun chantier enregistré dans cet espace.')
    with b:
        st.session_state['_kl_workspace']=st.session_state['kl_workspace']
        st.selectbox('Espace de travail',['Mes projets','Exemple fictif'],key='_kl_workspace',on_change=change_workspace)
    with c:st.button('Charger l’exemple fictif',on_click=load_example,icon=':material/science:',width='stretch',wrap=True,key='load_demo')
    if notice:=st.session_state.pop('kl_flash',None):st.success(notice)
    page=st.session_state['kl_page']
    pages={'Tableau de bord':dashboard,'Métrés et devis':measure_page,'Risques et données':risk_page,
           'Encyclopédie BTP':encyclopedia_page,'Paiements':payments_page,'Carnet de chantier':notebook_page,'Données et aide':data_page}
    pages[page](data,p)
    st.html('<footer class="kl-footer">KATIKA DATA LAB · Prototype pédagogique · Paiements simulés · Hypothèses et sources consultables dans chaque outil</footer>')


if __name__=='__main__':main()
