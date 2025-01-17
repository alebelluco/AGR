# QUESTA è LA VERSIONE ATTUALMENTE DISPONIBILE IN STREAMLIT

import streamlit  as st 
import pandas as pd 
import numpy as np 
from utils import utility as ut
import plotly_express as px
import plotly.graph_objects as go

import datetime as dt
from datetime import timedelta

st.set_page_config(layout='wide')
sx, dx = st.columns([4,1])
with sx:
    st.title('Arti Grafiche Reggiane e LAI')
    st.subheader('OEE Dashboard :red[v1_17.01.2025]')
with dx:
    st.image('https://github.com/alebelluco/AGR/blob/main/logo.png?raw=True')
st.divider()

import_config = {

    'STAM' : {
        'name' : 'CALENDARIO STAMPA', # Nome del foglio nel file excel
        'colonne' : 16, # numero di colonne da mantenere nel file
        'durata' : {'T1':8, 'T2':8, 'T3':8}, # durata di ogni turno
        'col_select' : [ #colonne da mantenere, vengono escluse le colonne consuntivo
                        "Giorno",
                        "Data",
                        "Turno",
                        "A121_cons_K08",
                        "A123_cons_105",
                        "A124_cons_162",
                        "A127_cons_145",
                        ]
    },
    'FUST' : {
        'name' : 'CALENDARIO FUSTELLA',
        'colonne' : 22,
        'durata' : {'T1':8, 'T2':8, 'T3':8},
        'col_select' : [
                        "Giorno",
                        "Data",
                        "Turno",
                        "A145_cons_CER",
                        "A148_cons_M1",
                        "A149_cons_M2",
                        "A152_cons_M3",
                        "A154_cons_M4",
                        "A155_cons_M5",
                        ]
    },
    'INCO' : {
        'name' : 'CALENDARIO PIEGA E INCOLLA',
        'colonne' : 30,
        'durata' : {'T1':8, 'T2':8, 'T3':8},
        'col_select' : [
                        "Giorno",
                        "Data",
                        "Turno",
                        "A164_cons_CTPK1",
                        "A161_cons_CTPK2",
                        "A165_cons_CTPK3",
                        "A174_cons_GBOX1",
                        "A175_cons_GBOX2",
                        "A172_cons_ALPINA110",
                        "A173_cons_ALPINA130",
                        "A151_cons_FINESTRA",
                        
                        ]
    }
}

path3 = st.sidebar.file_uploader('Caricare calendario')
if not path3:
    st.stop()

cal = ut.cal_upload(path3, import_config)
cal['variable'] = cal['variable']+ ' '

# SELEZIONE REPARTO E MACCHINA
selsx, selcx, seldx = st.columns([1,1,3])
with selsx:
    reparto = st.selectbox('Selezionare il reparto', options=['STAM','FUST','INCO'])
with selcx:
    placeholder = st.empty()


cal = cal[cal.reparto == reparto]

path = st.sidebar.file_uploader('Caricare file "KPI Indicatori Lean"')
if not path:
    st.stop()

v_tgt = ut.upload(path, 'VI_LEAN_INDIC_TOT_SAS')
v_tgt = v_tgt[v_tgt.PROGR_COMMESSA.astype(str)!='nan']

v_tgt['PROGR_COMMESSA'] = [int(comm) for comm in v_tgt.PROGR_COMMESSA]
v_tgt['PROGR_COMMESSA'] = [stringa[:4] for stringa in v_tgt['PROGR_COMMESSA'].astype(str)]
v_tgt['ANNO_COMMESSA'] =  [stringa[:2] for stringa in v_tgt['ANNO_COMMESSA'].astype(str)]

for i in range(len(v_tgt)):
    comm = v_tgt['PROGR_COMMESSA'].iloc[i]
    if len(comm) == 1:
        v_tgt['PROGR_COMMESSA'].iloc[i] = '000'+v_tgt['PROGR_COMMESSA'].iloc[i]
    elif len(comm) == 2:
        v_tgt['PROGR_COMMESSA'].iloc[i] = '00'+v_tgt['PROGR_COMMESSA'].iloc[i]
    elif len(comm) == 3:
        v_tgt['PROGR_COMMESSA'].iloc[i] = '0'+v_tgt['PROGR_COMMESSA'].iloc[i]
    else:
        v_tgt['PROGR_COMMESSA'].iloc[i] = v_tgt['PROGR_COMMESSA'].iloc[i]

#v_tgt = v_tgt[["PROGR_COMMESSA","COD_MACCHINA","VEL_TIRATURA_PREVISTA",]]
v_tgt['commessa'] = v_tgt['ANNO_COMMESSA'].astype(str)+ '00' + v_tgt['PROGR_COMMESSA']
v_tgt['key'] = (v_tgt['COD_MACCHINA'].astype(str) + v_tgt['commessa']).str.replace(" ","")
v_tgt = v_tgt[["key","VEL_TIRATURA_PREVISTA"]]
#v_tgt = v_tgt[["key","VEL_TIRATURA_PREVISTA"]]
v_tgt = v_tgt.drop_duplicates()


path2 = st.sidebar.file_uploader('Caricare file TBM')
if not path2:
    st.stop()

tbm = ut.upload(path2, 'Base_Dati')


tbm = tbm[tbm.DES_REPARTO == reparto]
tbm = tbm[tbm.COD_MACCHINA.astype(str) != 'A153 '] # la macchina è a villavara

#tbm = tbm[tbm.COD_MACCHINA == 'A124 ']

cols = [
  'inizio',
  'fine',
  "COD_MACCHINA",
  "DES_CLIENTE",
  "COD_COMMESSA",
  "CAPOCONTO",
  "DES_ARTICOLO",
  "NUM_FUSTELLA",
  "TIPO_ATTIVITA",
  "COD_DES_ATTIVITA",
  "RAGG_ATTIVITA",
  "DURATA_STEP",
  "QTA_PRODOTTA",
  "QTA_SCARTI",
  "DES_OPERATORE",
  "DES_REPARTO",
  "DES_OPERATORI",
  't1'
]

cols_test = [
  'inizio',
  'fine',
  'COD_COMMESSA',
  "DES_REPARTO",
  'COD_MACCHINA',
  "TIPO_ATTIVITA",
  "COD_DES_ATTIVITA",
  "RAGG_ATTIVITA",
  "DURATA_STEP",
  'data',
  'check',
  'turno',
  'pezzi_adj',
  'scarti_adj',
]

orari_turni = {
    'STAM':{
    't1':{'inizio':dt.time(5,30), 'fine':dt.time(13,30)},
    't2':{'inizio':dt.time(13,30), 'fine':dt.time(21,30)},
    't3':{'inizio':dt.time(21,30), 'fine':dt.time(5,30)}},
    'FUST':{
    't1':{'inizio':dt.time(5,30), 'fine':dt.time(13,30)},
    't2':{'inizio':dt.time(13,30), 'fine':dt.time(21,30)},
    't3':{'inizio':dt.time(21,30), 'fine':dt.time(5,30)}},
    'INCO':{
    't1':{'inizio':dt.time(5,30), 'fine':dt.time(13,30)},
    't2':{'inizio':dt.time(13,30), 'fine':dt.time(21,30)},
    't3':{'inizio':dt.time(21,30), 'fine':dt.time(5,30)}},
    'MVAR':{
    't1':{'inizio':dt.time(5,30), 'fine':dt.time(13,30)},
    't2':{'inizio':dt.time(13,30), 'fine':dt.time(21,30)},
    't3':{'inizio':dt.time(21,30), 'fine':dt.time(5,30)}},
    'CERN':{
    't1':{'inizio':dt.time(5,30), 'fine':dt.time(13,30)},
    't2':{'inizio':dt.time(13,30), 'fine':dt.time(21,30)},
    't3':{'inizio':dt.time(21,30), 'fine':dt.time(5,30)}},
    'ACCO':{
    't1':{'inizio':dt.time(5,30), 'fine':dt.time(13,30)},
    't2':{'inizio':dt.time(13,30), 'fine':dt.time(21,30)},
    't3':{'inizio':dt.time(21,30), 'fine':dt.time(5,30)}},
    'FINE':{
    't1':{'inizio':dt.time(5,30), 'fine':dt.time(13,30)},
    't2':{'inizio':dt.time(13,30), 'fine':dt.time(21,30)},
    't3':{'inizio':dt.time(21,30), 'fine':dt.time(5,30)}},
}

# preparazione timestamp

#tbm = tbm[['DATA_ORA']]
tbm['anno'] = [stringa[:4] for stringa in tbm.DATA_ORA]
tbm['mese'] = [stringa[4:6] for stringa in tbm.DATA_ORA]
tbm['giorno'] = [stringa[6:8] for stringa in tbm.DATA_ORA]
tbm['ora'] = [stringa[9:11] for stringa in tbm.DATA_ORA]
tbm['minuto'] = [stringa[12:14] for stringa in tbm.DATA_ORA]

tbm['anno'] = tbm['anno'].astype(int)
tbm['mese'] = tbm['mese'].astype(int)
tbm['giorno'] = tbm['giorno'].astype(int)
tbm['ora'] = tbm['ora'].astype(int)
tbm['minuto'] = tbm['minuto'].astype(int)

tbm['inizio'] = [dt.datetime(tbm.anno.iloc[i], tbm.mese.iloc[i], tbm.giorno.iloc[i], tbm.ora.iloc[i], tbm.minuto.iloc[i]) for i in range(len(tbm))]
tbm['fine'] = [tbm.inizio.iloc[i] + timedelta(hours=tbm.DURATA_STEP.iloc[i]) for i in range(len(tbm))]

tbm = ut.sdoppia(tbm,orari_turni,'inizio','fine',['COD_MACCHINA','inizio'])
tbm['data'] = [data.date() for data in tbm.inizio]

# attribuzione del turno corretto
tbm['turno']=None
for i in range(len(tbm)):
    reparto = tbm['DES_REPARTO'].iloc[i]
    turni = orari_turni[reparto]
    start = tbm['inizio'].iloc[i]
    
    if (start >= dt.datetime.combine(start.date(), turni['t1']['inizio'])) & (start < dt.datetime.combine(start.date(), turni['t1']['fine'])):
        turno = 1
    elif (start >= dt.datetime.combine(start.date(), turni['t2']['inizio'])) & (start < dt.datetime.combine(start.date(), turni['t2']['fine'])):
        turno = 2
    else:
        turno = 3
        if (start.time() >= dt.time(0,0)) & (start.time() < dt.time(21,0,0)):
            tbm['data'].iloc[i] = tbm['data'].iloc[i] - dt.timedelta(days=1)

    tbm['turno'].iloc[i] = turno


tbm['durata_calcolata'] = [np.round((tbm['fine'].iloc[i] - tbm['inizio'].iloc[i]).seconds/3600,5) for i in range(len(tbm))]
tbm['appoggio'] = np.where(tbm['DURATA_STEP'] != 0, tbm['DURATA_STEP'], 1)
tbm['appoggio_calc'] = np.where(tbm['durata_calcolata'] != 0, tbm['durata_calcolata'], 1)
tbm['ripartizione'] = np.where(tbm.check.astype(str) != 'None',tbm['appoggio_calc'] / tbm['appoggio'], 1)
tbm['pezzi_adj'] = tbm['QTA_PRODOTTA'] * tbm['ripartizione']
tbm['scarti_adj'] = tbm['QTA_SCARTI'] * tbm['ripartizione']
tbm['key'] = (tbm['COD_MACCHINA'] + tbm['COD_COMMESSA']).str.replace(" ","")
tbm = tbm.merge(v_tgt, how='left', left_on='key', right_on='key')

tbm = tbm[tbm.VEL_TIRATURA_PREVISTA.astype(str) != 'nan']
tbm = tbm[tbm.VEL_TIRATURA_PREVISTA != 0]
#st.write(tbm)

col_ragg = [
    'data',
    'turno',
    'COD_MACCHINA',
    #'check',
    'COD_COMMESSA',
    'durata_calcolata',
    'VEL_TIRATURA_PREVISTA',
    'pezzi_tot',
    'tc_std',
    #'velocità'
]


df_vel = tbm[tbm.RAGG_ATTIVITA == 'PRODUZIONE']
df_vel = df_vel[df_vel['durata_calcolata'] !=0 ]
df_vel['pezzi_tot'] = df_vel['pezzi_adj'] + df_vel['scarti_adj']

df_vel['tc_std'] = [1/vm for vm in df_vel['VEL_TIRATURA_PREVISTA']]

df_vel_agg = df_vel[col_ragg]   

#st.write(cal)
#st.write(df_vel)
#st.write(df_vel_agg)
#st.stop()


df_vel_agg = df_vel.groupby(by=['data','COD_MACCHINA', 'turno','COD_COMMESSA'], as_index=False).agg({'durata_calcolata': 'sum', 'VEL_TIRATURA_PREVISTA': 'mean', 'pezzi_tot':'sum', 'tc_std':'mean' })

df_vel_agg['velocità'] = (df_vel_agg['tc_std']*df_vel_agg['pezzi_tot'])/df_vel_agg['durata_calcolata']
df_vel_agg['key_agg'] = df_vel_agg['data'].astype(str) + df_vel_agg['turno'].astype(str)



# SELEZIONE MACCHINE
with placeholder:
    macchina_select = st.multiselect('Selezionare macchina', options = df_vel_agg.COD_MACCHINA.unique())

dic_reparti = {'STAM':'Stampa', 'FUST':'Fustella', 'INCO':'Piega Incolla'}
st.subheader(f'KPI reparto {dic_reparti[reparto]}', divider='grey')

if len(macchina_select) == 0:
    macchina_select = df_vel_agg.COD_MACCHINA.unique()

df_vel_agg = df_vel_agg[[any(macchina in check for macchina in macchina_select) for check in df_vel_agg.COD_MACCHINA.astype(str) ]]
df_vel = df_vel[[any(macchina in check for macchina in macchina_select) for check in df_vel.COD_MACCHINA.astype(str) ]]
cal = cal[[any(macchina in check for macchina in macchina_select) for check in cal.variable.astype(str)]]



df_scarti = tbm[['data','turno','pezzi_adj','scarti_adj']].copy() # include righe prod e non prod
df_scarti['pezzi_tot'] = tbm['pezzi_adj']+tbm['scarti_adj']

df_scarti['key'] = df_scarti['data'].astype(str) + df_scarti['turno'].astype(str)
df_scarti = df_scarti.groupby(by=['data','turno','key'], as_index=False).sum()
df_scarti = df_scarti[df_scarti.pezzi_tot != 0]

db_oee = cal[['Data','Turno','ttd','key']].copy() # non ho messo il reparto perchè devre essere filtrato a monte per tutto
db_oee = db_oee.groupby(by=['Data','Turno','key'], as_index=False).sum()

df_vel_agg_sum = df_vel_agg[['data','turno','durata_calcolata']].groupby(by=['data','turno'],as_index=False).sum()
df_vel_agg_sum['key'] = df_vel_agg_sum['data'].astype(str) + df_vel_agg_sum['turno'].astype(str)


db_oee = db_oee.merge(df_vel_agg_sum[['key','durata_calcolata']], how='left', left_on='key', right_on='key')
df_vel_agg = df_vel_agg.merge(df_vel_agg_sum[['key','durata_calcolata']], how='left', left_on='key_agg', right_on='key')

#l'aggregazione viene fatta senza la macchina, eventualmente viene filtrata prima

df_vel_agg['quota_v'] = df_vel_agg['durata_calcolata_x'] / df_vel_agg['durata_calcolata_y']
df_vel_agg['v_pesata'] = df_vel_agg['velocità']*df_vel_agg['quota_v']

vel_output = df_vel_agg[['data','turno','v_pesata']].groupby(by=['data','turno'], as_index=False).sum()

vel_output['turno'] = vel_output['turno'].astype(int)
vel_output = vel_output.sort_values(by=['data','turno'], ascending=True)
vel_output['key'] = vel_output['data'].astype(str)+vel_output['turno'].astype(str)
vel_output['key_graph'] = vel_output['data'].astype(str)+' | '+vel_output['turno'].astype(str)


db_oee = db_oee.merge(vel_output[['key','v_pesata']], how='left', left_on='key',right_on='key').fillna(0)
db_oee = db_oee.merge(df_scarti[['key','pezzi_tot','scarti_adj']], how='left', left_on='key',right_on='key').dropna()

db_oee = db_oee[db_oee.durata_calcolata!=0]

db_oee['key_graph']=db_oee['Data'].astype(str)+' | '+db_oee['Turno'].astype(str)
db_oee['Disponibilità'] = db_oee['durata_calcolata'] / db_oee['ttd']
db_oee['Qualità'] = (db_oee['pezzi_tot'] - db_oee['scarti_adj'])/db_oee['pezzi_tot']
db_oee['OEE'] = db_oee['Disponibilità'] * db_oee['v_pesata'] * db_oee['Qualità']

date_range = st.slider('sleziona intervallo date', db_oee.Data.min(), db_oee.Data.max(), (db_oee.Data.min(), db_oee.Data.max()))
db_oee = db_oee[(db_oee.Data > date_range[0]) & (db_oee.Data < date_range[1]) ]

# valori medi
# i turni con ttd 0 non pesano sulla media
db_oee['oee_medio'] = db_oee[db_oee.ttd != 0]['OEE'].mean()
db_oee['d_media'] = db_oee[db_oee.ttd != 0]['Disponibilità'].mean()
db_oee['v_media'] = db_oee[db_oee.ttd != 0]['v_pesata'].mean()
db_oee['q_media'] = db_oee[db_oee.ttd != 0]['Qualità'].mean()

oee = go.Figure()
dettaglio = go.Figure()

oee.add_trace(go.Scatter(
    x=db_oee['key_graph'],
    y=db_oee['OEE']
))


f_size = 15
oee.update_layout(
        yaxis=dict(
            title=dict(text="OEE", font = dict(size=f_size)),
            side="left",
            range=[0, 1.1],
            tickformat=".0%",
            tickfont=dict(size=f_size)))

dettaglio.add_trace(go.Bar(
    x= db_oee['key_graph'],
    y= db_oee['Disponibilità'],
    name='D',
))

dettaglio.add_trace(go.Bar(
    x= db_oee['key_graph'],
    y= db_oee['v_pesata'],
    name = 'V'
))

dettaglio.add_trace(go.Bar(
    x= db_oee['key_graph'],
    y= db_oee['Qualità'],
    name='Q'
))

dettaglio.update_layout(
        yaxis=dict(
            title=dict(text="PCT", font = dict(size=f_size)),
            side="left",
            range=[0, 1.1],
            tickformat=".0%",
            tickfont=dict(size=f_size)))

st.subheader('Andamento OEE')
st.plotly_chart(oee, use_container_width=True)

st.divider()
st.subheader('OEE medio nel periodo selezionato: {:0.2f}%'.format(db_oee.oee_medio.iloc[0]*100))
st.write('Disponibilità media: {:0.2f}%'.format(db_oee.d_media.iloc[0]*100))
st.write('Velocità media: {:0.2f}%'.format(db_oee.v_media.iloc[0]*100))
st.write('Qualità media: {:0.2f}%'.format(db_oee.q_media.iloc[0]*100))
st.divider()
st.subheader('Dettglio componenti OEE')
st.plotly_chart(dettaglio, use_container_width=True)

# ANALISI DI DETTAGLIO 


if st.toggle('drill'):
    data_select = st.selectbox('data', options = cal.Data.unique())
    turno_select = st.selectbox('turno', options = [1,2,3])
    df_vel = df_vel[(df_vel.turno == turno_select) & (df_vel.data == data_select)]
    pezzi = df_vel.pezzi_tot.sum()
    durata = df_vel.durata_calcolata.sum()
    st.dataframe(df_vel[['data','turno','inizio','fine','COD_MACCHINA','COD_COMMESSA','RAGG_ATTIVITA','durata_calcolata','pezzi_tot','VEL_TIRATURA_PREVISTA']],use_container_width=True)
    #st.subheader('OEE', divider='red')
    db_oee_print = db_oee[(db_oee.Turno.astype(int) == turno_select) & (db_oee.Data == data_select)]
    st.dataframe(db_oee_print, use_container_width=True)

tbm['key_disp'] = tbm['data'].astype(str) + tbm['COD_MACCHINA'] + '-'+tbm.turno.astype(str)
tbm_disp = tbm[tbm.RAGG_ATTIVITA != 'PRODUZIONE'].copy()


# Note
# Ci sono righe che non trovano il valore della velocità target perchè non sono ancora chiuse nel file KPI (solitamente è riferito all'ultimo gg o in corso)
# Tali righe sono eliminate insieme a quelle con v_tgt = 0 (ininfluenti)

