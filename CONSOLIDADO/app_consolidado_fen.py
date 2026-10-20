"""Dashboard consolidado UIR + UAT · FEN 2026–2027.
Ejecutar: streamlit run app_consolidado_fen.py
Archivos: EXCEL DS 164 EF(2).xlsx, datos_fen_2026(1).xlsx,
plantilla_avance_fen_2026.xlsx, coordenadas_distritos.xlsx.
"""
from pathlib import Path
import re
import hashlib
import unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Consolidado FEN 2026–2027", page_icon="🇵🇪", layout="wide")
BASE = Path(__file__).resolve().parent
COLORES = {"UIR": "#1976D2", "UAT": "#16A34A"}
ESTADOS = {"CONTRATADO":"#0891B2", "EN CONTRATACION":"#F59E0B", "EN EJECUCION":"#16A34A", "CULMINADO":"#2563EB", "ATENDIDO EN OTROS PROYECTOS":"#94A3B8"}
SEMAFORO = [(100,"Cumplido (≥100%)","#15803D"),(70,"Aceptable (70–99%)","#84CC16"),(40,"En riesgo (40–69%)","#F59E0B"),(0,"Crítico (<40%)","#DC2626")]
NIVELES = ["DEPARTAMENTO","PROVINCIA","DISTRITO"]
DEP_COORD = {"PIURA":(-5.1945,-80.6328),"LA LIBERTAD":(-8.1116,-79.0287),"ANCASH":(-9.5278,-77.5278),"LIMA":(-12.0464,-77.0428),"AREQUIPA":(-16.409,-71.5375)}
KIT_CLAVES = [("ALIMENTO SUPLEMENTARIO","Alimento suplementario"),("INSUMOS AGRICOLAS","Insumos agrícolas"),("VETERINARIO","Veterinario"),("RIEGO","Riego"),("SEMILLAS","Semillas"),("PRADERAS","Praderas"),("ABREVADERO","Abrevadero")]

st.markdown('''<style>.stApp{background:#F2F8F4}.block-container{max-width:1550px;padding-top:1.5rem}.hero{padding:23px 29px;background:linear-gradient(110deg,#14532D,#15803D);border-radius:15px;color:white;margin-bottom:20px}.hero h1{color:white;margin:0;font-size:29px}.hero p{margin:7px 0 0;color:#DCFCE7}.section{border-left:5px solid #16A34A;padding-left:12px;font-size:21px;font-weight:750;color:#14532D;margin:24px 0 12px}</style>''',unsafe_allow_html=True)

def norm(x):
    if pd.isna(x): return ""
    s = unicodedata.normalize("NFKD",str(x).strip().upper())
    return " ".join("".join(ch for ch in s if not unicodedata.combining(ch)).split())

def fichero(nombre,alternativas):
    for n in alternativas:
        if (BASE/n).exists(): return BASE/n
    st.error(f"Falta {nombre}. Se buscó: {', '.join(alternativas)}")
    st.stop()

def nombre_kit(x):
    n=norm(x)
    for k,et in KIT_CLAVES:
        if k in n:return et
    return str(x).strip().title()

def numero(x):return pd.to_numeric(x,errors="coerce").fillna(0)

def estado_pct(x):
    if pd.isna(x):return "Sin avance reportado"
    return next(et for limite,et,_ in SEMAFORO if x>=limite)

@st.cache_data(show_spinner="Leyendo información de UIR y UAT…")
def leer(f_uir,f_uat,f_av,f_co,firmas):
    u=pd.read_excel(f_uir)
    u.columns=[str(c).strip() for c in u.columns]
    u=u.rename(columns={"Ben":"BENEFICIARIOS","Ha":"HECTAREAS","Km":"KM"})
    for c in NIVELES:
        orig=c.title()
        u[c]=u[orig].map(norm)
    u["ESTADO"]=u["Estado"].map(norm).replace({"CONTRADO":"CONTRATADO","DESESTIMADA":"ATENDIDO EN OTROS PROYECTOS"})
    u["PAQUETE"]=u["Paquete"].fillna("").map(norm).str.replace(r"\s*PQ\.?\s*(\d+)$",r" PQ \1",regex=True)
    u.loc[~u["PAQUETE"].str.contains(r"PQ \d+$",regex=True),"PAQUETE"]="SIN PAQUETE"
    for c in ("BENEFICIARIOS","HECTAREAS","KM"):u[c]=numero(u[c])
    u.loc[u["ESTADO"].eq("ATENDIDO EN OTROS PROYECTOS"),["BENEFICIARIOS","HECTAREAS","KM"]]=0
    u["UNIDAD"]="UIR"
    u["SECTOR"]=u["Sector"].fillna("").astype(str) if "Sector" in u else ""

    a=pd.read_excel(f_uat,sheet_name="Hoja1",header=None)
    h=next((i for i in range(len(a)) if norm(a.iat[i,1])=="DEPARTAMENTO"),None)
    if h is None:raise ValueError("UAT: no se encontró la fila de encabezado DEPARTAMENTO")
    filas=a.iloc[h+1:].copy()
    filas=filas[filas[1].notna() & ~filas[0].astype(str).str.upper().str.contains("TOTAL",na=False)]
    partes=[]
    for j in range(4,a.shape[1]-1,2):
        etiqueta=a.iat[h-1,j]
        if pd.isna(etiqueta):continue
        temp=pd.DataFrame({"DEPARTAMENTO":filas[1].map(norm),"PROVINCIA":filas[2].map(norm),"DISTRITO":filas[3].map(norm),"KIT":nombre_kit(etiqueta),"META":numero(filas[j]),"BENEF_EXCEL":numero(filas[j+1])})
        partes.append(temp)
    t=pd.concat(partes,ignore_index=True)
    t["PROG"]=t["BENEF_EXCEL"]
    t.loc[t["KIT"].eq("Praderas"),"PROG"]=t.loc[t["KIT"].eq("Praderas"),"META"]/5*3
    t=t[(t["META"]>0)|(t["PROG"]>0)].copy()
    t["AV_META"]=0.0;t["AVANCE"]=0.0
    t["UNIDAD"]="UAT"
    tiene_av=False;sin_coincidencia=0
    if f_av is not None:
        av=pd.read_excel(f_av)
        av.columns=[str(c).strip().upper() for c in av.columns]
        if all(c in av for c in NIVELES+["KIT"]) and ("AVANCE_META" in av or "AVANCE_BENEFICIARIOS" in av):
            tiene_av=True
            for c in NIVELES:av[c]=av[c].map(norm)
            av["KIT"]=av["KIT"].map(nombre_kit)
            for c in ("AVANCE_META","AVANCE_BENEFICIARIOS"):
                if c not in av:av[c]=0
                av[c]=numero(av[c])
            keys=NIVELES+["KIT"]
            g=av.groupby(keys,as_index=False)[["AVANCE_META","AVANCE_BENEFICIARIOS"]].sum()
            sin_coincidencia=len(g.merge(t[keys].drop_duplicates(),on=keys,how="left",indicator=True).query("_merge == 'left_only'"))
            t=t.drop(columns="AV_META").merge(g,on=keys,how="left")
            t["AV_META"]=t["AVANCE_META"].fillna(0)
            t["AVANCE"]=t["AVANCE_BENEFICIARIOS"].fillna(0)
            pr=t["KIT"].eq("Praderas")
            t.loc[pr,"AVANCE"]=t.loc[pr,"AV_META"]/5*3
            t=t.drop(columns=["AVANCE_META","AVANCE_BENEFICIARIOS"])
    t["PCT"]=np.where(t["PROG"]>0,t["AVANCE"]/t["PROG"]*100,np.nan)

    coord=pd.DataFrame(columns=NIVELES+["LAT","LON","APROXIMADA"])
    if f_co is not None:
        coord=pd.read_excel(f_co)
        for c in NIVELES:coord[c]=coord[c].map(norm)
        coord=coord.drop_duplicates(NIVELES)
    return u,t,coord,tiene_av,sin_coincidencia

f_uir=fichero("Excel UIR",["EXCEL DS 164 EF(2).xlsx","EXCEL DS 164 EF.xlsx"])
f_uat=fichero("Excel UAT",["datos_fen_2026(1).xlsx","datos_fen_2026.xlsx"])
f_av=next((BASE/n for n in ["plantilla_avance_fen_2026.xlsx","avance_fen_2026.xlsx"] if (BASE/n).exists()),None)
f_co=(BASE/"coordenadas_distritos.xlsx") if (BASE/"coordenadas_distritos.xlsx").exists() else None
archivos=[p for p in (f_uir,f_uat,f_av,f_co) if p is not None]
try:
    uir,uat,coord,HAY_AV,SIN_COINCIDENCIA=leer(f_uir,f_uat,f_av,f_co,tuple((str(p),p.stat().st_mtime_ns) for p in archivos))
except Exception as exc:
    st.exception(exc);st.stop()


# DISEÑO EJECUTIVO V2: una sola narrativa visual, con detalles bajo pestañas.
st.markdown("""<style>
.block-container{max-width:1460px;padding:1.1rem 2.1rem 2rem}
.hero{margin-bottom:12px;padding:23px 28px;border-radius:18px;background:linear-gradient(105deg,#123F2A,#198754)}
.hero h1{font-size:28px;letter-spacing:-.5px}.hero p{font-size:13px}
div[data-testid="stMetric"]{background:white;border:1px solid #D8E7DD;border-radius:13px;padding:13px 17px;min-height:107px;box-shadow:0 2px 8px #164E2710}
div[data-testid="stMetricLabel"]{color:#52685C;font-size:13px}
div[data-testid="stMetricValue"]{color:#174B32;font-weight:750}
div[data-testid="stVerticalBlockBorderWrapper"]{border-color:#D8E7DD!important;border-radius:15px;background:white}
div[data-testid="stPlotlyChart"]{border-radius:12px;overflow:hidden}
div[data-testid="stTabs"] button{font-weight:700}
.section{font-size:18px;margin:23px 0 12px;border-left:4px solid #198754}
[data-testid="stCaptionContainer"]{color:#607568}
</style>""",unsafe_allow_html=True)

def seccion(t): st.markdown(f'<div class="section">{t}</div>',unsafe_allow_html=True)
def graficar(fig,key):
    fig.update_layout(template="plotly_white",paper_bgcolor="white",plot_bgcolor="white",margin=dict(l=15,r=15,t=55,b=18),font=dict(color="#274737",size=12),legend_title=None)
    st.plotly_chart(fig,use_container_width=True,key=key)
def tarjeta_vacia(msg): st.info(msg)

st.markdown('<div class="hero"><h1>🌿 Dashboard ejecutivo | FEN 2026–2027</h1><p>Consolidado territorial de intervenciones · UIR / Descolmatación &nbsp;·&nbsp; UAT / Intervenciones agropecuarias</p></div>',unsafe_allow_html=True)

geo=pd.concat([uir[NIVELES],uat[NIVELES]],ignore_index=True).drop_duplicates()
geo=geo[(geo[NIVELES]!="").all(axis=1)]
seccion('01 · Filtros del territorio')
f1,f2,f3,f4=st.columns([1.15,1.15,1.15,1])
dep=f1.selectbox('Departamento',['Todo el Perú']+sorted(geo.DEPARTAMENTO.unique()),key='v2_dep')
g=geo if dep=='Todo el Perú' else geo[geo.DEPARTAMENTO.eq(dep)]
prov=f2.selectbox('Provincia',['Todas las provincias']+sorted(g.PROVINCIA.unique()),key='v2_prov')
g=g if prov=='Todas las provincias' else g[g.PROVINCIA.eq(prov)]
dist=f3.selectbox('Distrito',['Todos los distritos']+sorted(g.DISTRITO.unique()),key='v2_dist')
vista=f4.selectbox('Unidad',['Ambas unidades','UIR','UAT'],key='v2_unidad')

def filtrar(d):
    if dep!='Todo el Perú': d=d[d.DEPARTAMENTO.eq(dep)]
    if prov!='Todas las provincias': d=d[d.PROVINCIA.eq(prov)]
    if dist!='Todos los distritos': d=d[d.DISTRITO.eq(dist)]
    return d
u=filtrar(uir) if vista!='UAT' else uir.iloc[:0]
t=filtrar(uat) if vista!='UIR' else uat.iloc[:0]
if u.empty and t.empty:
    st.warning('No hay registros para estos filtros.');st.stop()
loc_u=u[NIVELES].drop_duplicates();loc_t=t[NIVELES].drop_duplicates()
territorios=pd.concat([loc_u,loc_t],ignore_index=True).drop_duplicates()
compartidos=loc_u.merge(loc_t,on=NIVELES,how='inner')

seccion('02 · Panorama consolidado')
k=st.columns(5)
k[0].metric('Departamentos',territorios.DEPARTAMENTO.nunique())
k[1].metric('Provincias',len(territorios[NIVELES[:2]].drop_duplicates()))
k[2].metric('Distritos únicos',len(territorios))
k[3].metric('Intervenciones UIR',len(u))
k[4].metric('Registros UAT',len(t))
st.caption(f'Cobertura compartida: {len(compartidos)} distritos presentes en ambas unidades. Los distritos se contabilizan sin duplicados.')

# Comparación simple: cifras distintas en tarjetas, no en un gráfico de suma.
a,b=st.columns(2)
with a:
    with st.container(border=True):
        st.markdown('**🌊 UIR · Descolmatación**')
        c1,c2,c3=st.columns(3)
        c1.metric('Paquetes',u.loc[u.PAQUETE.ne('SIN PAQUETE'),'PAQUETE'].nunique())
        c2.metric('Beneficiarios',f'{u.BENEFICIARIOS.sum():,.0f}')
        c3.metric('Kilómetros',f'{u.KM.sum():,.2f}')
with b:
    with st.container(border=True):
        st.markdown('**🌱 UAT · Intervenciones agropecuarias**')
        c1,c2,c3=st.columns(3)
        c1.metric('Tipos de intervención',t.KIT.nunique())
        c2.metric('Atenciones',f'{t.PROG.sum():,.0f}')
        c3.metric('Distritos',len(loc_t))

seccion('03 · Explorador territorial interactivo')
st.caption('Un mapa compartido, con UIR en azul y UAT en verde. Haz clic en un punto (o en una barra) para bajar de departamento a provincia y a distrito.')

# Cada nivel territorial agrupa las ubicaciones de cada unidad, para evitar
# mostrar centenares de distritos en la vista nacional.
if 'ruta_mapa_v5' not in st.session_state:
    st.session_state.ruta_mapa_v5 = []
firma = (dep, prov, dist, vista)
if st.session_state.get('firma_mapa_v5') != firma:
    st.session_state.ruta_mapa_v5 = []
    st.session_state.firma_mapa_v5 = firma

bases=[]
for unidad, frame in [('UIR',u),('UAT',t)]:
    if frame.empty: continue
    if unidad=='UIR':
        detalle=frame[NIVELES].copy()
        detalle['VALOR']=frame['BENEFICIARIOS'].to_numpy()
    else:
        detalle=frame[NIVELES].copy()
        detalle['VALOR']=frame['PROG'].to_numpy()
    detalle['UNIDAD']=unidad
    bases.append(detalle)
base_mapa=pd.concat(bases,ignore_index=True) if bases else pd.DataFrame(columns=NIVELES+['VALOR','UNIDAD'])

# ── CORRECCIÓN DEL MAPA (1/3): coordenadas con cadena de respaldo ──────────────
# distrito (exacta) → centroide de su provincia → centroide de su departamento
# → capital departamental. Así los puntos de UIR sin coordenada propia dejan de
# caer todos en la capital del departamento. Lo no exacto se marca COORD_APROX.
if not coord.empty and {'LAT','LON'}<=set(coord.columns):
    coords=coord[NIVELES+['LAT','LON']].drop_duplicates(NIVELES).dropna(subset=['LAT','LON'])
    base_mapa=base_mapa.merge(coords,on=NIVELES,how='left')
    base_mapa['COORD_APROX']=base_mapa['LAT'].isna() | base_mapa['LON'].isna()
    prov_c=(coords.groupby(['DEPARTAMENTO','PROVINCIA'],as_index=False)[['LAT','LON']].mean()
            .rename(columns={'LAT':'LAT_P','LON':'LON_P'}))
    dep_c=(coords.groupby('DEPARTAMENTO',as_index=False)[['LAT','LON']].mean()
           .rename(columns={'LAT':'LAT_D','LON':'LON_D'}))
    base_mapa=(base_mapa.merge(prov_c,on=['DEPARTAMENTO','PROVINCIA'],how='left')
                        .merge(dep_c,on='DEPARTAMENTO',how='left'))
    for c_lat,c_lon in (('LAT_P','LON_P'),('LAT_D','LON_D')):
        base_mapa['LAT']=base_mapa['LAT'].fillna(base_mapa[c_lat])
        base_mapa['LON']=base_mapa['LON'].fillna(base_mapa[c_lon])
    base_mapa=base_mapa.drop(columns=['LAT_P','LON_P','LAT_D','LON_D'])
else:
    base_mapa['LAT']=np.nan
    base_mapa['LON']=np.nan
    base_mapa['COORD_APROX']=True
for nombre,(lat,lon) in DEP_COORD.items():
    mascara=base_mapa.DEPARTAMENTO.eq(nombre) & base_mapa.LAT.isna()
    base_mapa.loc[mascara,'LAT']=lat
    base_mapa.loc[mascara,'LON']=lon

ruta=st.session_state.ruta_mapa_v5
for i, lugar in enumerate(ruta):
    candidatos=base_mapa.copy()
    for j, anterior in enumerate(ruta[:i]):
        candidatos=candidatos[candidatos[NIVELES[j]].eq(anterior)]
    if lugar not in candidatos[NIVELES[i]].unique():
        st.session_state.ruta_mapa_v5=ruta[:i]
        break
ruta=st.session_state.ruta_mapa_v5


sc=base_mapa.copy()
for i, lugar in enumerate(ruta):
    sc=sc[sc[NIVELES[i]].eq(lugar)]
nivel=min(len(ruta),2)
columna=NIVELES[nivel]
agrupacion=['UNIDAD']+NIVELES[:nivel+1]
mapa_data=(sc.groupby(agrupacion,as_index=False)
           .agg(VALOR=('VALOR','sum'),LAT=('LAT','mean'),LON=('LON','mean'),
                COORD_APROX=('COORD_APROX','all')))
mapa_data['LUGAR']=mapa_data[columna]
mapa_data['NOMBRE']=mapa_data['LUGAR'].str.title()

# El tamaño de los puntos se calcula INDEPENDIENTEMENTE por unidad.
# Si se comparte la escala de beneficiarios, los puntos de UIR desaparecen
# visualmente frente a los valores mucho mayores de UAT.
mapa_data['TAM']=18.0
for unidad in ('UIR','UAT'):
    mascara=mapa_data.UNIDAD.eq(unidad)
    if mascara.any():
        valores=mapa_data.loc[mascara,'VALOR'].clip(lower=0)
        vmax=valores.max()
        mapa_data.loc[mascara,'TAM']=14+18*np.sqrt(valores/vmax) if vmax>0 else 18

# La referencia nacional UIR es el departamento (no el promedio de los
# distritos UAT que casualmente coinciden en el archivo de coordenadas).
if len(ruta)==0:
    for dep_nombre,(lat,lon) in DEP_COORD.items():
        mascara=mapa_data.UNIDAD.eq('UIR') & mapa_data.DEPARTAMENTO.eq(dep_nombre)
        mapa_data.loc[mascara,'LAT']=lat
        mapa_data.loc[mascara,'LON']=lon
        mapa_data.loc[mascara,'COORD_APROX']=True

# ── CORRECCIÓN DEL MAPA (2/3): abrir en círculo los puntos que comparten coordenada ──
# Si varias provincias/distritos de una misma unidad caen en el mismo punto
# (p. ej. UIR sin coordenada propia), quedan apilados y no se puede hacer clic en
# cada uno. Se reparten en círculo alrededor de ese punto (posición aproximada).
def dispersar(df,base):
    df=df.sort_values(['UNIDAD','LUGAR']).copy()
    ok=df['LAT'].notna() & df['LON'].notna()
    clave=df['UNIDAD']+'|'+df['LAT'].round(4).astype(str)+'|'+df['LON'].round(4).astype(str)
    for _,idx in df[ok].groupby(clave[ok]).groups.items():
        n=len(idx)
        if n>1:
            ang=np.linspace(0,2*np.pi,n,endpoint=False)
            radio=base*(1+n/8)
            df.loc[idx,'LAT']=df.loc[idx,'LAT'].to_numpy()+radio*np.sin(ang)
            df.loc[idx,'LON']=df.loc[idx,'LON'].to_numpy()+radio*np.cos(ang)
            df.loc[idx,'COORD_APROX']=True
    return df
mapa_data=dispersar(mapa_data,{0:.25,1:.12,2:.04}[nivel])

# Desplazamiento visual solo para puntos de ambas unidades que se solapan.
mapa_data['LON_VIS']=mapa_data['LON'].copy()
for lugar,grupo in mapa_data.groupby('LUGAR'):
    if grupo.UNIDAD.nunique()>1:
        mapa_data.loc[grupo.index[grupo.UNIDAD.eq('UIR')],'LON_VIS']-=.16 if len(ruta)==0 else .035
        mapa_data.loc[grupo.index[grupo.UNIDAD.eq('UAT')],'LON_VIS']+=.16 if len(ruta)==0 else .035

# Navegación territorial: selector junto a los botones superiores.
enc1,enc_ir,enc2,enc3=st.columns([3.0,2.7,1.25,1.25],vertical_alignment='center')
enc1.markdown('<div style="padding-top:8px;font-weight:700">🇵🇪 '+ ' › '.join(['Perú']+[x.title() for x in ruta])+'</div>',unsafe_allow_html=True)
if len(ruta)<3 and not mapa_data.empty:
    opciones=sorted(mapa_data.LUGAR.dropna().unique().tolist())
    elegir=enc_ir.selectbox(f'Ir a {columna.lower()}', ['Seleccionar…']+opciones,
        key='ir_v5_'+str(firma)+'_'+str(ruta), label_visibility='collapsed',
        help=f'Selecciona un {columna.lower()} para profundizar en el mapa')
    if elegir!='Seleccionar…':
        st.session_state.ruta_mapa_v5=ruta+[elegir];st.rerun()
else:
    enc_ir.empty()
if enc2.button('⬆ Subir',disabled=not ruta,key='subir_v5',use_container_width=True):
    st.session_state.ruta_mapa_v5=ruta[:-1];st.rerun()
if enc3.button('🏠 Perú',disabled=not ruta,key='peru_v5',use_container_width=True):
    st.session_state.ruta_mapa_v5=[];st.rerun()

# ── CORRECCIÓN DEL MAPA (3/3): identificar el lugar clicado de forma robusta ──
# Primero por customdata; si el navegador no lo devuelve, se reconstruye con
# el número de curva (una por unidad) y el índice del punto.
def lugar_clic(pt,dfp):
    cd=pt.get('customdata')
    if cd:
        return cd[0]
    try:
        unidades=[x for x in ('UAT','UIR') if x in set(dfp.UNIDAD.astype(str))]
        sub=dfp[dfp.UNIDAD.astype(str).eq(unidades[pt['curve_number']])].reset_index(drop=True)
        i=pt.get('point_index',pt.get('point_number'))
        return sub.iloc[i].LUGAR
    except Exception:
        return None

izq,der=st.columns([1,1],gap='medium')
with izq:
    with st.container(border=True):
        st.markdown('**Cobertura territorial · UIR y UAT**')
        puntos=mapa_data.dropna(subset=['LAT','LON_VIS']).copy()
        puntos['UNIDAD']=pd.Categorical(puntos['UNIDAD'],categories=['UAT','UIR'],ordered=True)
        puntos=puntos.sort_values('UNIDAD')
        if puntos.empty:
            st.info('No se encontraron coordenadas para este territorio.')
        else:
            rango=max(float(puntos.LAT.max()-puntos.LAT.min()),float(puntos.LON.max()-puntos.LON.min())*.7,.15)
            zoom=float(np.clip(np.log2(600/rango),4,10))
            fig=px.scatter_map(puntos,lat='LAT',lon='LON_VIS',color='UNIDAD',
                color_discrete_map=COLORES,size='TAM',size_max=35,
                category_orders={'UNIDAD':['UAT','UIR']},
                hover_name='NOMBRE',custom_data=['LUGAR','UNIDAD'],
                hover_data={'VALOR':':,.0f','TAM':False,'LAT':False,'LON_VIS':False,'COORD_APROX':True},
                map_style='carto-positron',height=570,
                center={'lat':float(puntos.LAT.mean()),'lon':float(puntos.LON.mean())},zoom=zoom)
            fig.update_traces(marker=dict(opacity=.88))
            fig.update_layout(margin=dict(l=0,r=0,t=0,b=0),
                legend=dict(orientation='h',x=.02,y=.02,title=None,bgcolor='rgba(255,255,255,.85)'))
            try:
                evento=st.plotly_chart(fig,use_container_width=True,
                    key='mapa_v5_'+str(firma)+'_'+str(ruta),on_select='rerun',selection_mode='points')
                pts=evento.selection.points
                if pts and len(ruta)<3:
                    lug=lugar_clic(pts[0],puntos)
                    if lug and lug in mapa_data.LUGAR.values:
                        st.session_state.ruta_mapa_v5=ruta+[lug];st.rerun()
            except TypeError:
                st.plotly_chart(fig,use_container_width=True)

with der:
    with st.container(border=True):
        st.markdown(f'**Beneficiarios / atenciones por {columna.lower()}**')
        ranking=mapa_data.sort_values('VALOR',ascending=False).copy()
        ranking['ETIQUETA']=ranking['LUGAR'].str.title()
        if len(ruta)==0:
            ranking=ranking.head(28)
        elif len(ranking)>30:
            ranking=ranking.head(30)
        ranking=ranking.sort_values('VALOR')
        if ranking.empty:
            st.info('Sin datos para mostrar.')
        else:
            graf=px.bar(ranking,y='ETIQUETA',x='VALOR',color='UNIDAD',
                orientation='h',barmode='group',color_discrete_map=COLORES,
                custom_data=['LUGAR'],height=570,text_auto=',.0f',
                labels={'VALOR':'Beneficiarios / atenciones','ETIQUETA':''})
            graf.update_layout(margin=dict(l=8,r=12,t=20,b=10),legend=dict(orientation='h',y=1.06),paper_bgcolor='white',plot_bgcolor='white')
            graf.update_xaxes(rangemode='tozero')
            if ranking.VALOR.sum()==0:
                st.info('Este territorio no registra beneficiarios/atenciones en los datos (sectores con 0 beneficiarios o intervenciones atendidas en otros proyectos). Revisa el detalle más abajo.')
            try:
                ev_b=st.plotly_chart(graf,use_container_width=True,
                    key='barras_v5_'+str(firma)+'_'+str(ruta),on_select='rerun',selection_mode='points')
                pts_b=ev_b.selection.points
                if pts_b and len(ruta)<3:
                    cd=pts_b[0].get('customdata')
                    if cd and cd[0] in mapa_data.LUGAR.values:
                        st.session_state.ruta_mapa_v5=ruta+[cd[0]];st.rerun()
            except TypeError:
                st.plotly_chart(graf,use_container_width=True)

# ── DETALLE DEL TERRITORIO SELECCIONADO ─────────────────────────────────────────
# Al bajar por el mapa (departamento › provincia › distrito) se muestra la ficha
# del lugar elegido: cifras clave y las tablas de intervenciones UIR y UAT.
def por_ruta(d):
    for i,l in enumerate(ruta):
        d=d[d[NIVELES[i]].eq(l)]
    return d

if ruta:
    d_u=por_ruta(u); d_t=por_ruta(t)
    with st.container(border=True):
        st.markdown('**📍 Detalle de '+' › '.join(x.title() for x in ruta)+'**')
        cu,ct=st.columns(2,gap='medium')
        with cu:
            st.markdown('🌊 **UIR · Descolmatación**')
            if d_u.empty:
                st.caption('Sin intervenciones UIR en este territorio.')
            else:
                m=st.columns(4)
                m[0].metric('Intervenciones',len(d_u))
                m[1].metric('Beneficiarios',f'{d_u.BENEFICIARIOS.sum():,.0f}')
                m[2].metric('Hectáreas',f'{d_u.HECTAREAS.sum():,.2f}')
                m[3].metric('Km',f'{d_u.KM.sum():,.2f}')
                st.caption(' · '.join(f'{k.title()}: {v}' for k,v in d_u.ESTADO.value_counts().items()))
        with ct:
            st.markdown('🌱 **UAT · Intervenciones agropecuarias**')
            if d_t.empty:
                st.caption('Sin intervenciones UAT en este territorio.')
            else:
                m=st.columns(3)
                m[0].metric('Tipos de intervención',d_t.KIT.nunique())
                m[1].metric('Atenciones',f'{d_t.PROG.sum():,.0f}')
                m[2].metric('Cumplimiento',f'{d_t.AVANCE.sum()/d_t.PROG.sum()*100:.1f}%' if HAY_AV and d_t.PROG.sum()>0 else 'Sin avance')
        extra=['DISTRITO'] if len(ruta)<3 else []
        if not d_u.empty:
            with st.expander('Ver intervenciones UIR (sectores)',expanded=len(ruta)==3):
                st.dataframe(d_u[extra+['SECTOR','PAQUETE','ESTADO','BENEFICIARIOS','HECTAREAS','KM']],
                             hide_index=True,use_container_width=True)
        if not d_t.empty:
            with st.expander('Ver intervenciones UAT (kits)',expanded=len(ruta)==3):
                cols_t=extra+['KIT','META','PROG']+(['AV_META','AVANCE','PCT'] if HAY_AV else [])
                st.dataframe(d_t[cols_t],hide_index=True,use_container_width=True,
                             column_config={'PCT':st.column_config.ProgressColumn('Cumplimiento',min_value=0,max_value=100,format='%.0f%%')})
else:
    st.caption('🔵 UIR · Descolmatación  |  🟢 UAT · Intervenciones.')
    st.caption('Tamaños normalizados por unidad. Haz clic en un punto o en una barra para ver el detalle. Las ubicaciones aproximadas (sin coordenada exacta de distrito) no representan coordenadas exactas. Navega por departamento, provincia y distrito.')

# ═════════════ 04 · INTEGRACIÓN UIR + UAT ═════════════
seccion('05 · Resultados por unidad')
st.info('**Lectura del cruce:** en esta selección UIR y UAT no comparten ningún distrito. '
                'Por eso no tiene sentido calcular una correlación estadística de beneficiarios entre ambas unidades.')
tab1=st.container();tab2=st.container()
with tab1:
    st.markdown('#### 🌊 UIR · Descolmatación')
    if u.empty:st.info('UIR no tiene registros en la selección actual.')
    else:
        p1,p2=st.columns(2)
        paquete=p1.selectbox('Paquete',['Todos los paquetes']+sorted(u.PAQUETE.unique()),key='v2_pq')
        est=p2.selectbox('Estado',['Todos los estados']+sorted(u.ESTADO.unique()),key='v2_est')
        du=u.copy()
        if paquete!='Todos los paquetes':du=du[du.PAQUETE.eq(paquete)]
        if est!='Todos los estados':du=du[du.ESTADO.eq(est)]
        ks=st.columns(5)
        ks[0].metric('Intervenciones',len(du))
        ks[1].metric('En contratación',du.ESTADO.eq('EN CONTRATACION').sum())
        ks[2].metric('En ejecución',du.ESTADO.eq('EN EJECUCION').sum())
        ks[3].metric('Culminadas',du.ESTADO.eq('CULMINADO').sum())
        ks[4].metric('Hectáreas',f'{du.HECTAREAS.sum():,.2f}')
        q1,q2=st.columns(2)
        with q1:
            estados=du.ESTADO.value_counts().rename_axis('Estado').reset_index(name='Intervenciones')
            if len(estados):
                fig=px.pie(estados,names='Estado',values='Intervenciones',hole=.56,color='Estado',color_discrete_map=ESTADOS,title='Situación de las intervenciones')
                fig.update_layout(height=360);graficar(fig,'v2_uir_dona')
        with q2:
            gp=du[du.PAQUETE.ne('SIN PAQUETE')].groupby('PAQUETE').size().reset_index(name='Intervenciones').sort_values('Intervenciones',ascending=False).head(12)
            if len(gp):
                fig=px.bar(gp.sort_values('Intervenciones'),x='Intervenciones',y='PAQUETE',orientation='h',color_discrete_sequence=['#17805A'],title='Paquetes con más intervenciones')
                fig.update_layout(height=360,yaxis_title=None);graficar(fig,'v2_uir_paquetes')
        with st.expander('Ver tabla de intervenciones UIR y descargar'):
            cols=NIVELES+['SECTOR','PAQUETE','ESTADO','BENEFICIARIOS','HECTAREAS','KM']
            st.dataframe(du[cols],hide_index=True,use_container_width=True)
            st.download_button('Descargar UIR (CSV)',du[cols].to_csv(index=False).encode('utf-8-sig'),'uir_filtrado.csv','text/csv')
with tab2:
    st.markdown('#### 🌱 UAT · Intervenciones y metas')
    if t.empty:st.info('UAT no tiene registros en la selección actual.')
    else:
        kit=st.selectbox('Intervención',['Todas las intervenciones']+sorted(t.KIT.unique()),key='v2_kit')
        dt=t if kit=='Todas las intervenciones' else t[t.KIT.eq(kit)]
        ks=st.columns(4)
        ks[0].metric('Tipos de intervención',dt.KIT.nunique())
        ks[1].metric('Metas / registros',len(dt))
        ks[2].metric('Atenciones programadas',f'{dt.PROG.sum():,.0f}')
        ks[3].metric('Cumplimiento',f'{dt.AVANCE.sum()/dt.PROG.sum()*100:.1f}%' if HAY_AV and dt.PROG.sum()>0 else 'Sin avance')
        g=dt.groupby('KIT',as_index=False)[['PROG','AVANCE']].sum()
        q1,q2=st.columns(2)
        with q1:
            fig=px.pie(g,names='KIT',values='PROG',hole=.56,title='Atenciones programadas por intervención')
            fig.update_layout(height=365);graficar(fig,'v2_uat_dona')
        with q2:
            if HAY_AV:
                largo=g.melt(id_vars='KIT',value_vars=['PROG','AVANCE'],var_name='Tipo',value_name='Atenciones')
                fig=px.bar(largo,y='KIT',x='Atenciones',color='Tipo',orientation='h',barmode='group',color_discrete_map={'PROG':'#BAE6FD','AVANCE':'#16A34A'},title='Programado frente a avance')
            else:fig=px.bar(g,y='KIT',x='PROG',orientation='h',title='Programación por intervención',color_discrete_sequence=['#16A34A'])
            fig.update_layout(height=365,yaxis_title=None);graficar(fig,'v2_uat_barras')
        if HAY_AV:
            sem=dt.groupby(NIVELES,as_index=False)[['PROG','AVANCE']].sum()
            sem['Porcentaje']=np.where(sem.PROG>0,sem.AVANCE/sem.PROG*100,np.nan)
            sem['Estado']=sem.Porcentaje.map(estado_pct)
            s=sem.Estado.value_counts().rename_axis('Estado').reset_index(name='Distritos')
            fig=px.bar(s,x='Estado',y='Distritos',color='Estado',color_discrete_map={x[1]:x[2] for x in SEMAFORO},title='Semáforo de cumplimiento por distrito')
            fig.update_layout(height=290,showlegend=False,xaxis_title=None);graficar(fig,'v2_semaforo')
        st.markdown('**Leyenda de cumplimiento:** 🟢 Cumplido (≥100%) · 🟩 Aceptable (70–99%) · 🟠 En riesgo (40–69%) · 🔴 Crítico (<40%)')
        with st.expander('Ver tabla UAT y descargar'):
            cols=NIVELES+['KIT','META','PROG']+(['AV_META','AVANCE','PCT'] if HAY_AV else [])
            st.dataframe(dt[cols],hide_index=True,use_container_width=True)
            st.download_button('Descargar UAT (CSV)',dt[cols].to_csv(index=False).encode('utf-8-sig'),'uat_filtrado.csv','text/csv')

with st.expander('Notas metodológicas y calidad de datos'):
    st.markdown('**UIR:** se excluye del impacto el estado «Atendido en otros proyectos». **UAT:** Praderas = meta física ÷ 5 × 3. **Cobertura:** distritos únicos entre unidades. **Beneficiarios:** no se suman UIR y UAT, pues no se ha validado que sean personas únicas ni indicadores equivalentes.')
    if SIN_COINCIDENCIA:st.warning(f'{SIN_COINCIDENCIA} combinaciones de avance UAT no coinciden con la programación.')
st.caption('Fuentes: UIR · EXCEL DS 164 EF | UAT · datos_fen_2026, plantilla de avance y coordenadas distritales.')
