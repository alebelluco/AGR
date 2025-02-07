# QUESTA è LA VERSIONE ATTUALMENTE DISPONIBILE IN STREAMLIT

import streamlit  as st 
import pandas as pd 
import numpy as np 
from utils import utility as ut
import plotly_express as px
import plotly.graph_objects as go

import datetime as dt
from datetime import timedelta
from utils import grafici as gh

st.set_page_config(layout='wide')
sx, dx = st.columns([4,1])
with sx:
    st.title('Arti Grafiche Reggiane e LAI')
    st.subheader('OEE Dashboard :red[v2_7.02.2025]')
with dx:
    st.image('https://github.com/alebelluco/AGR/blob/main/logo.png?raw=True')
st.divider()
istruzioni = st.expander('Help')
with istruzioni:
    st.write('Cliccare la freccetta in alto a sinistra per aprire la barra laterale;')
    st.write('Nel primo riquadro caricare il file "CALENDARIO FABBRICA_WEEK.xlsx"')
    st.write('Nel secondo riquadro caricare il file "KPI - Indicatori_LEAN_SAS.xlsx"')
    st.write('Nel secondo riquadro caricare il file "estrazione ultimi 30 giorni.xlsx"')


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
                        "A170_cons_CODIFICA",
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
selsx, selcx, seldx = st.columns([1,1,2])
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

if reparto=='INCO':
    v_tgt = v_tgt.drop_duplicates()

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
tbm = tbm[tbm.COD_MACCHINA.astype(str) != 'A176 ']


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

# calcolo del ttd per piegaincolla

ttd_inco = tbm.copy()
ttd_inco = ttd_inco.drop(columns='VEL_TIRATURA_PREVISTA')
ttd_inco = ttd_inco.drop_duplicates()
ttd_inco['chiave'] = ttd_inco.COD_MACCHINA + ttd_inco['data'].astype(str) + ttd_inco.turno.astype(str)
ttd_inco['durata_calcolata']=ttd_inco['durata_calcolata']*ttd_inco['ripartizione']
ttd_inco = ttd_inco[['chiave','COD_MACCHINA','durata_calcolata']]
ttd_inco = ttd_inco.groupby(by=['COD_MACCHINA','chiave'], as_index=False).sum()


if reparto == 'INCO':
    cal['chiave']=cal['variable'] + cal['key']
    cal = cal.drop(columns='ttd')
    cal = cal.merge(ttd_inco, how='left', left_on='chiave', right_on='chiave')
    cal = cal.rename(columns={'durata_calcolata':'ttd'})
    cal = cal[cal.ttd.astype(str) != 'nan']

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

df_vel_agg = df_vel.groupby(by=['data','COD_MACCHINA', 'turno','COD_COMMESSA'], as_index=False).agg({'durata_calcolata': 'sum', 'VEL_TIRATURA_PREVISTA': 'mean', 'pezzi_tot':'sum', 'tc_std':'mean' })

df_vel_agg['velocità'] = (df_vel_agg['tc_std']*df_vel_agg['pezzi_tot'])/df_vel_agg['durata_calcolata']
df_vel_agg['key_agg'] = df_vel_agg['data'].astype(str) + df_vel_agg['turno'].astype(str)


# SELEZIONE MACCHINE
with placeholder:
    macchina_select = st.multiselect('Selezionare macchina (nessuna selezione = tutto il reparto)', options = df_vel_agg.COD_MACCHINA.unique())

dic_reparti = {'STAM':'Stampa', 'FUST':'Fustella', 'INCO':'Piega Incolla'}
st.subheader(f'KPI reparto {dic_reparti[reparto]}', divider='grey')

if len(macchina_select) == 0:
    macchina_select = df_vel_agg.COD_MACCHINA.unique()

df_vel_agg = df_vel_agg[[any(macchina in check for macchina in macchina_select) for check in df_vel_agg.COD_MACCHINA.astype(str) ]]
df_vel = df_vel[[any(macchina in check for macchina in macchina_select) for check in df_vel.COD_MACCHINA.astype(str) ]]
cal = cal[[any(macchina in check for macchina in macchina_select) for check in cal.variable.astype(str)]]

tbm = tbm[[any(macchina in check for macchina in macchina_select) for check in tbm.COD_MACCHINA.astype(str) ]]
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

if reparto == 'INCO':
    db_oee = db_oee[db_oee.Turno != '3']

# valori medi
# i turni con ttd 0 non pesano sulla media
db_oee['oee_medio'] = db_oee[db_oee.ttd != 0]['OEE'].mean()
db_oee['d_media'] = db_oee[db_oee.ttd != 0]['Disponibilità'].mean()
db_oee['v_media'] = db_oee[db_oee.ttd != 0]['v_pesata'].mean()
db_oee['q_media'] = db_oee[db_oee.ttd != 0]['Qualità'].mean()
db_oee['SMA_15'] = db_oee.OEE.rolling(15).mean()
db_oee['SMA_45'] = db_oee.OEE.rolling(45).mean()

oee = go.Figure()


oee.add_trace(go.Scatter(
    x=db_oee['key_graph'],
    y=db_oee['OEE']
))

f_size = 15
oee.update_layout(
        yaxis=dict(
            title=dict(text="OEE", font = dict(size=f_size)),
            side="left",
            range=[0, db_oee.OEE.max()*1.1],
            tickformat=".0%",
            tickfont=dict(size=f_size)))


sma = go.Figure()

sma.add_trace(go.Scatter(
    x=db_oee['key_graph'],
    y=db_oee['SMA_15'],
    line=dict(color='red'),
    name='SMA_15'
))

sma.add_trace(go.Scatter(
    x=db_oee['key_graph'],
    y=db_oee['SMA_45'],
    line=dict(color='white'),
    name='SMA_45'
))

sma.update_layout(
        yaxis=dict(
            range=[0, db_oee.SMA_15.max() * 1.02],
            tickformat=".0%",
            tickfont=dict(size=f_size)))




dettaglio = go.Figure()


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
            range=[0, db_oee['v_pesata'].max()*1.2],
            tickformat=".0%",
            tickfont=dict(size=f_size)))

sx_oee, dx_oee = st.columns([2,1])

with sx_oee:

    st.subheader('Andamento OEE')
    st.plotly_chart(oee, use_container_width=True)

with dx_oee:
    st.subheader('Andamento Medie mobili OEE')
    st.plotly_chart(sma, use_container_width=True)


st.divider()

st.subheader('OEE medio nel periodo selezionato: {:0.2f}%'.format(db_oee.oee_medio.iloc[0]*100))
st.write('Disponibilità media: {:0.2f}%'.format(db_oee.d_media.iloc[0]*100))
st.write('Velocità media: {:0.2f}%'.format(db_oee.v_media.iloc[0]*100))
st.write('Qualità media: {:0.2f}%'.format(db_oee.q_media.iloc[0]*100))
st.divider()

tab1, tab2, tab3, tab4 = st.tabs(['Dettaglio componenti','Analisi velocità', 'Perdite di disponibilità', 'TBM'])
with tab1:



    st.subheader('Dettglio componenti OEE')
    st.plotly_chart(dettaglio, use_container_width=True)



    tbm['key_disp'] = tbm['data'].astype(str) + tbm['COD_MACCHINA'] + ' | '+tbm.turno.astype(str)
    tbm_disp = tbm[tbm.RAGG_ATTIVITA != 'PRODUZIONE'].copy()
    tbm_disp = tbm_disp[(tbm_disp.data > date_range[0]) & (tbm_disp.data < date_range[1]) ]

    perdite_macchina = tbm_disp[['COD_MACCHINA','durata_calcolata']]
    perdite_causale = tbm_disp[['COD_DES_ATTIVITA','durata_calcolata']]

    
    #st.write(tbm_disp)

    # ANALISI DI DETTAGLIO 

    if st.toggle('Visualizza dettagli turno'):
        data_select = st.selectbox('data', options = cal.Data.unique())
        turno_select = st.selectbox('turno', options = [1,2,3])
        df_vel = df_vel[(df_vel.turno == turno_select) & (df_vel.data == data_select)]
        pezzi = df_vel.pezzi_tot.sum()
        durata = df_vel.durata_calcolata.sum()
        st.dataframe(df_vel[['data','turno','inizio','fine','COD_MACCHINA','COD_COMMESSA','RAGG_ATTIVITA','durata_calcolata','pezzi_tot','VEL_TIRATURA_PREVISTA']],use_container_width=True)
        #st.subheader('OEE', divider='red')
        db_oee_print = db_oee[(db_oee.Turno.astype(int) == turno_select) & (db_oee.Data == data_select)]
        st.dataframe(db_oee_print, use_container_width=True)


with tab2:

    # analisi delle velocità 
    # raggruppamento sulle commesse
    df_vel_focus = df_vel[['COD_MACCHINA','COD_COMMESSA','CAPOCONTO','durata_calcolata','VEL_TIRATURA_PREVISTA','pezzi_tot','tc_std']]

    df_vel_focus = df_vel_focus.groupby(by=['COD_MACCHINA','COD_COMMESSA','CAPOCONTO','VEL_TIRATURA_PREVISTA','tc_std'], as_index=False).sum()
    df_vel_focus['ton'] = df_vel_focus['tc_std'] * df_vel_focus['pezzi_tot']
    df_vel_focus['vel'] = df_vel_focus['ton'] / df_vel_focus['durata_calcolata']
    df_vel_focus['v_reale'] = df_vel_focus['pezzi_tot'] / df_vel_focus['durata_calcolata']
    df_vel_focus=df_vel_focus[(df_vel_focus.vel <= 3 )]
    sx_2, dx_2 = st.columns([1,1])
    with sx_2:
        st.subheader('Distribuzione delle velocità', divider='grey')
        h_v = go.Figure(data=[go.Histogram(x=df_vel_focus.vel, histnorm='probability', marker_color='grey')])
        h_v.add_vrect(x0=0, x1=0.3, line_width=0, fillcolor="red", opacity=0.2)
        h_v.add_vrect(x0=1, x1=df_vel_focus.vel.max()+0.1, line_width=0, fillcolor="red", opacity=0.2)

        st.plotly_chart(h_v, use_container_width=True)

    with dx_2:
        st.subheader('Dettaglio velocità fuori limite', divider='grey')
        out = df_vel_focus[(df_vel_focus.vel >= 1 )| (df_vel_focus.vel <= 0.3)].sort_values(by='vel', ascending=False).reset_index(drop=True)
        st.dataframe(out)
        ut.scarica_excel(out,'v_da_analizzare.xlsx')

with tab3:

    f_size = 18
    f_angle =-45

    stile = {
        'colore_barre':'#D9D9D9',
        'colore_linea':'#CD3128',
        'name_bar':'Durata',
        'name_cum':'cum_pct',
        'y_name': 'Durata [h]',
        'y2_name': 'pct_cumulativa',
        'tick_size':16,
        'angle':-45
    }


    sx3 ,dx3 = st.columns([1,2])
    with sx3:
        st.subheader('Perdite per macchina')
        st.divider()
        #perdite_macchina 
        pareto_macchine = gh.pareto(perdite_macchina,'COD_MACCHINA','durata_calcolata',stile)
        pareto_macchine.update_layout(height=570)

        st.plotly_chart(pareto_macchine, use_container_width=True)
        macc_aggr = perdite_macchina.groupby(by='COD_MACCHINA').sum().sort_values(by='durata_calcolata', ascending=False)
        #st.dataframe(macc_aggr)

    with dx3:

        st.subheader('Perdite per causale')
        st.divider()
        #perdite_causale 
        pareto_causali = gh.pareto(perdite_causale,'COD_DES_ATTIVITA','durata_calcolata',stile)
        pareto_causali.update_layout(height=700)
        sx33,dx33 = st.columns([3,2])
        with sx33:
            st.plotly_chart(pareto_causali, use_container_width=True)
        
        with dx33:
            st.write('Dettaglio')
            st.divider()
            caus_aggr = perdite_causale.groupby(by='COD_DES_ATTIVITA').sum().sort_values(by='durata_calcolata', ascending=False)
            st.dataframe(caus_aggr, use_container_width=True) 

with tab4:
    top_5 = list(caus_aggr.index[:5])

    top_sel = st.multiselect('Selezionare causale (nessuna selezione = tutte le causali)', options=top_5)
    if len(top_sel) == 0:
        top_sel = top_5

    print_rank = [
  "DATA_ORA",
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
  "VEL_TIRATURA_PREVISTA"
]
    rank = tbm_disp[[any(causa in check for causa in top_sel) for check in tbm_disp.COD_DES_ATTIVITA]].sort_values(by='durata_calcolata', ascending=False)
    st.dataframe(rank[print_rank].reset_index(drop=True))

    pass


        


    # Note
    # Ci sono righe che non trovano il valore della velocità target perchè non sono ancora chiuse nel file KPI (solitamente è riferito all'ultimo gg o in corso)
    # Tali righe sono eliminate insieme a quelle con v_tgt = 0 (ininfluenti)
    # 7/02/2025 piega incolla, vengono trovati più valori di velocità target nel file kpi lean --> manenere solo il massimo oppure fare la media prima di fare il merge

    # Appunti
    #h_v.add_vline(x=1, line_width=2, line_dash="dash", line_color="red")
