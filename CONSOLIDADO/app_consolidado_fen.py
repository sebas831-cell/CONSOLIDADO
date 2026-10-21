"""Dashboard consolidado UIR + UAT · FEN 2026–2027.
Ejecutar: streamlit run app_consolidado_fen.py
Archivos: EXCEL DS 164 EF(2).xlsx, datos_fen_2026(1).xlsx,
plantilla_avance_fen_2026.xlsx, coordenadas_distritos.xlsx.
Todo el tema visual está incluido en este mismo archivo.
"""
from pathlib import Path
import re
import hashlib
import unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
try:
    from streamlit_plotly_events import plotly_events
except ImportError:
    plotly_events = None

# ═════════════ TEMA VISUAL (estilo, plantilla de gráficos y cabecera) ═════════════
FUENTE = "Manrope, system-ui, sans-serif"
COLORWAY = ["#1F8A55", "#1976D2", "#F59E0B", "#7C3AED", "#0891B2",
            "#DB2777", "#65A30D", "#64748B", "#EA580C", "#0F766E"]

PLANTILLA = go.layout.Template(layout=dict(
    font=dict(family=FUENTE, color="#274737", size=12),
    colorway=COLORWAY,
    paper_bgcolor="white",
    plot_bgcolor="white",
    title=dict(font=dict(size=15, color="#10281C"), x=0.01, xanchor="left"),
    xaxis=dict(gridcolor="#EAF1EC", zerolinecolor="#DCE8E0", linecolor="#DCE8E0", ticks=""),
    yaxis=dict(gridcolor="#EAF1EC", zerolinecolor="#DCE8E0", linecolor="#DCE8E0", ticks=""),
    hoverlabel=dict(bgcolor="white", bordercolor="#DCE8E0",
                    font=dict(family=FUENTE, color="#10281C", size=12)),
))

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
:root{--ink:#10281C;--muted:#5B6F63;--line:#DCE8E0;--bg:#F4F8F5;--brand:#14532D;--brand2:#1F8A55}
html,body,.stApp,[class*="css"]{font-family:'Manrope',system-ui,sans-serif}
.stApp{background:var(--bg);color:var(--ink)}
header[data-testid="stHeader"]{background:transparent}
#MainMenu,footer{visibility:hidden}
.block-container{max-width:1460px;padding:1.2rem 2.2rem 3rem}

/* Cabecera */
.hero{position:relative;overflow:hidden;margin-bottom:18px;padding:28px 32px;border-radius:20px;color:#fff;
  background:
   repeating-radial-gradient(circle at 94% -10%,transparent 0 26px,rgba(255,255,255,.055) 26px 27px),
   linear-gradient(115deg,#0E3B22 0%,#14532D 45%,#1F8A55 100%);
  box-shadow:0 10px 30px -12px rgba(20,83,45,.55)}
.hero h1{margin:0;color:#fff;font-size:30px;font-weight:800;letter-spacing:-.6px;line-height:1.15}
.hero p{margin:8px 0 0;color:#D6F5E2;font-size:14px;font-weight:500}
.hero .chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:16px}
.hero .chip{padding:5px 12px;border-radius:999px;font-size:12px;font-weight:700;
  background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.22)}

/* Títulos de sección */
.section{display:flex;align-items:center;gap:14px;margin:30px 0 14px;
  font-size:17px;font-weight:800;color:var(--brand);letter-spacing:-.2px;border:0;padding:0}
.section::after{content:"";flex:1;height:1px;background:var(--line)}

/* Indicadores */
div[data-testid="stMetric"]{background:#fff;border:1px solid var(--line);border-left:4px solid var(--brand2);
  border-radius:14px;padding:14px 18px;min-height:104px;box-shadow:0 6px 16px -10px rgba(20,83,45,.35)}
div[data-testid="stMetricLabel"]{color:var(--muted);font-size:12.5px;font-weight:600}
div[data-testid="stMetricValue"]{color:var(--brand);font-weight:800;font-size:28px;letter-spacing:-.5px}

/* Contenedores con borde */
div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stVerticalBlock"][class*="border"]{border-color:var(--line)!important;border-radius:16px;background:#fff}
div[data-testid="stPlotlyChart"]{border-radius:12px;overflow:hidden}
div[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:12px;overflow:hidden}

/* Controles */
div[data-baseweb="select"]>div{border-radius:10px;border-color:var(--line);background:#fff;min-height:42px}
div[data-baseweb="select"]>div:hover{border-color:var(--brand2)}
.stSelectbox label,.stRadio label p{font-weight:700;color:var(--ink)}
.stButton>button,.stDownloadButton>button{border-radius:10px;font-weight:700;border:1px solid var(--line);
  transition:box-shadow .15s,border-color .15s}
.stButton>button:hover{border-color:var(--brand2);box-shadow:0 4px 12px -6px rgba(20,83,45,.5);color:var(--brand)}
.stButton>button[kind="primary"]{background:var(--brand);border-color:var(--brand);color:#fff}

/* Pestañas y expanders */
div[data-testid="stTabs"] button{font-weight:700;color:var(--muted)}
div[data-testid="stTabs"] button[aria-selected="true"]{color:var(--brand)}
div[data-testid="stTabs"] [data-baseweb="tab-highlight"]{background:var(--brand2);height:3px}
details{border:1px solid var(--line)!important;border-radius:14px!important;background:#fff}
details summary{font-weight:700;color:var(--brand)}
[data-testid="stCaptionContainer"]{color:var(--muted)}
h4{color:var(--brand);font-weight:800;letter-spacing:-.3px}

/* Alertas del portafolio UIR */
.alert-red,.alert-amber,.alert-green{padding:14px 16px;border-radius:12px;height:100%;font-size:14px;line-height:1.5;
  border:1px solid transparent;border-left-width:5px}
.alert-red{background:#FEF2F2;border-color:#FECACA;border-left-color:#DC2626;color:#7F1D1D}
.alert-amber{background:#FFFBEB;border-color:#FDE68A;border-left-color:#F59E0B;color:#78350F}
.alert-green{background:#F0FDF4;border-color:#BBF7D0;border-left-color:#16A34A;color:#14532D}
</style>
"""


def aplicar_tema():
    """Inyecta fuente y CSS. Llamar una vez, justo después de st.set_page_config."""
    st.markdown(CSS, unsafe_allow_html=True)


def estilizar(fig):
    """Aplica la plantilla y márgenes comunes a cualquier figura Plotly."""
    fig.update_layout(template=PLANTILLA, margin=dict(l=15, r=15, t=55, b=18), legend_title=None)
    try:
        fig.update_layout(barcornerradius=4)  # esquinas redondeadas (Plotly >= 5.19)
    except Exception:
        pass
    return fig


def hero(titulo, subtitulo, chips=()):
    """Cabecera principal con etiquetas opcionales."""
    c = "".join(f'<span class="chip">{x}</span>' for x in chips)
    st.markdown(
        f'<div class="hero"><h1>{titulo}</h1><p>{subtitulo}</p>'
        + (f'<div class="chips">{c}</div>' if c else "") + "</div>",
        unsafe_allow_html=True)

st.set_page_config(page_title="Consolidado FEN 2026–2027", page_icon="🇵🇪", layout="wide")
aplicar_tema()
BASE = Path(__file__).resolve().parent
COLORES = {"UIR": "#1976D2", "UAT": "#16A34A"}
ESTADOS = {"CONTRATADO":"#0891B2", "EN CONTRATACION":"#F59E0B", "EN EJECUCION":"#16A34A", "CULMINADO":"#2563EB", "ATENDIDO EN OTROS PROYECTOS":"#94A3B8"}
SEMAFORO = [(100,"Cumplido (≥100%)","#15803D"),(70,"Aceptable (70–99%)","#84CC16"),(40,"En riesgo (40–69%)","#F59E0B"),(0,"Crítico (<40%)","#DC2626")]
NIVELES = ["DEPARTAMENTO","PROVINCIA","DISTRITO"]
DEP_COORD = {"PIURA":(-5.1945,-80.6328),"LA LIBERTAD":(-8.1116,-79.0287),"ANCASH":(-9.5278,-77.5278),"LIMA":(-12.0464,-77.0428),"AREQUIPA":(-16.409,-71.5375)}
KIT_CLAVES = [("ALIMENTO SUPLEMENTARIO","Alimento suplementario"),("INSUMOS AGRICOLAS","Insumos agrícolas"),("VETERINARIO","Veterinario"),("RIEGO","Riego"),("SEMILLAS","Semillas"),("PRADERAS","Praderas"),("ABREVADERO","Abrevadero")]

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


# DISEÑO EJECUTIVO: el estilo vive en el bloque TEMA VISUAL (arriba).
def seccion(t): st.markdown(f'<div class="section">{t}</div>',unsafe_allow_html=True)
def graficar(fig,key):
    estilizar(fig)
    st.plotly_chart(fig,use_container_width=True,key=key)
def tarjeta_vacia(msg): st.info(msg)

hero("🌿 Dashboard ejecutivo | FEN 2026–2027",
     "Consolidado territorial de intervenciones agropecuarias y de descolmatación",
     chips=["UIR · Descolmatación","UAT · Intervenciones agropecuarias","Cobertura nacional"])

seccion('01 · Selecciona una unidad para explorar sus intervenciones')
st.caption('Selecciona UIR o UAT en el círculo o mediante los botones. El mapa y todos los indicadores se actualizarán según la unidad seleccionada.')

def serie_avance_uir(d):
    """Serie de avance (0-100) por fila si el Excel UIR trae columna de avance; si no, None."""
    nombres=('AVANCE','% AVANCE','AVANCE %','PORCENTAJE AVANCE','AVANCE FISICO','AVANCE_FISICO_%','AVANCE FISICO %')
    for c in d.columns:
        if norm(c) in nombres:
            v=pd.to_numeric(d[c],errors='coerce')
            if v.notna().any():
                if v.max()<=1: v=v*100
                return v.clip(0,100)
    return None

def avance_uir(d):
    """Avance promedio de UIR; si el Excel no trae columna de avance, asume 100%."""
    s=serie_avance_uir(d)
    return float(s.mean()) if s is not None and s.notna().any() else float('nan')

def avance_uat(d):
    if HAY_AV and d.PROG.sum()>0:
        return float(np.clip(d.AVANCE.sum()/d.PROG.sum()*100,0,100))
    return float('nan')

unidades_av=[]
if not uir.empty: unidades_av.append(('UIR','🌊 UIR · Descolmatación',avance_uir(uir)))
if not uat.empty: unidades_av.append(('UAT','🌱 UAT · Intervenciones agropecuarias',avance_uat(uat)))

if 'v6_sel' not in st.session_state: st.session_state.v6_sel=None
if st.session_state.v6_sel not in [x[0] for x in unidades_av]:
    st.session_state.v6_sel=None

# El círculo SIEMPRE está completo: cada unidad ocupa una porción fija.
# El porcentaje de avance se muestra en la etiqueta, no cambia el tamaño.
fig_av=go.Figure(go.Pie(
    labels=[x[0] for x in unidades_av],
    values=[1]*len(unidades_av),
    hole=.58,sort=False,direction='clockwise',rotation=180,
    marker=dict(colors=[COLORES[x[0]] for x in unidades_av],line=dict(color='white',width=4)),
    text=[f'<b>{x[0]}</b><br>{x[2]:.0f}% avance' for x in unidades_av],
    textinfo='text',textfont=dict(size=17,color='white'),
    pull=[0.06 if x[0]==st.session_state.v6_sel else 0 for x in unidades_av],
    hovertemplate='<b>%{label}</b><br>Avance: %{customdata:.1f}%<extra></extra>',
    customdata=[x[2] for x in unidades_av]))
estilizar(fig_av)
fig_av.update_layout(height=380,showlegend=False,margin=dict(l=10,r=10,t=20,b=10),
    annotations=[dict(text='<b>Intervenciones</b><br>FEN 2026–2027',x=.5,y=.5,showarrow=False,font=dict(size=15,color='#174B32'))])

_,c_dona,_=st.columns([1,1.4,1])
with c_dona:
    with st.container(border=True):
        # Los gráficos Pie no emiten selección fiable mediante st.plotly_chart.
        # plotly_events captura clics reales sobre cada sector de la dona.
        if plotly_events is not None:
            # Cambiar la clave al cambiar de unidad evita reutilizar el clic anterior.
            clave_dona = 'dona_click_' + str(st.session_state.v6_sel or 'inicio')
            clics = plotly_events(fig_av, click_event=True, select_event=False,
                                  hover_event=False, override_height=380,
                                  key=clave_dona)
            if clics:
                idx = clics[0].get('pointNumber', clics[0].get('pointIndex'))
                if idx is not None and 0 <= int(idx) < len(unidades_av):
                    codigo = unidades_av[int(idx)][0]
                    st.session_state.v6_sel = None if st.session_state.v6_sel == codigo else codigo
                    st.rerun()
        else:
            st.plotly_chart(fig_av, use_container_width=True, key='v6_dona_visual')
            st.caption('Para activar el clic directamente en el círculo, instala: pip install streamlit-plotly-events')
        # Botones con el nombre de cada unidad (respaldo fiable al clic en el gráfico)
        bcols=st.columns(len(unidades_av))
        for col,(cod,nombre,pct) in zip(bcols,unidades_av):
            activo=st.session_state.v6_sel==cod
            if col.button(f'{nombre.split(" · ")[0]} · {pct:.0f}%' if np.isfinite(pct) else nombre.split(' · ')[0]+' · Sin avance',key=f'v6_btn_{cod}',
                          type='primary' if activo else 'secondary',use_container_width=True):
                st.session_state.v6_sel=None if activo else cod
                st.rerun()


if st.session_state.v6_sel is None:
    st.info('Selecciona UIR o UAT en el gráfico superior para ver sus indicadores, filtros, mapa y detalle.')
    st.stop()

# Los filtros se alimentan EXCLUSIVAMENTE de la unidad activa.
vista=st.session_state.v6_sel
geo=(uir if vista=='UIR' else uat)[NIVELES].drop_duplicates().copy()
geo=geo[(geo[NIVELES]!="").all(axis=1)]

# Al cambiar de unidad, limpiar filtros de la unidad anterior.
if st.session_state.get('filtros_unidad_anterior') != vista:
    for clave in ('v2_dep','v2_prov','v2_dist'):
        st.session_state.pop(clave,None)
    st.session_state.filtros_unidad_anterior=vista

seccion('02 · Filtros del territorio · '+vista)
f1,f2,f3,f4=st.columns([1.15,1.15,1.15,1])
ops_dep=['Todo el Perú']+sorted(geo.DEPARTAMENTO.unique())
if st.session_state.get('v2_dep') not in ops_dep:
    st.session_state.v2_dep='Todo el Perú'
dep=f1.selectbox('Departamento',ops_dep,key='v2_dep')
g=geo if dep=='Todo el Perú' else geo[geo.DEPARTAMENTO.eq(dep)]
ops_prov=['Todas las provincias']+sorted(g.PROVINCIA.unique())
if st.session_state.get('v2_prov') not in ops_prov:
    st.session_state.v2_prov='Todas las provincias'
prov=f2.selectbox('Provincia',ops_prov,key='v2_prov')
g=g if prov=='Todas las provincias' else g[g.PROVINCIA.eq(prov)]
ops_dist=['Todos los distritos']+sorted(g.DISTRITO.unique())
if st.session_state.get('v2_dist') not in ops_dist:
    st.session_state.v2_dist='Todos los distritos'
dist=f3.selectbox('Distrito',ops_dist,key='v2_dist')
vista=st.session_state.v6_sel or 'Ambas unidades'
f4.selectbox('Unidad activa', [vista], index=0, disabled=True, key='unidad_activa_visual')

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

# El panorama solo se despliega después de elegir una unidad.
seccion('03 · Panorama territorial · '+('UIR' if vista=='UIR' else 'UAT'))
if vista=='UIR':
    k=st.columns(3)
    k[0].metric('Departamentos',territorios.DEPARTAMENTO.nunique())
    k[1].metric('Provincias',len(territorios[NIVELES[:2]].drop_duplicates()))
    k[2].metric('Distritos',len(territorios))
    st.markdown('#### 📌 Indicadores ejecutivos · UIR')
    total=len(u)
    paquetes=u.loc[u.PAQUETE.ne('SIN PAQUETE'),'PAQUETE'].nunique()
    estados=u.ESTADO.value_counts()
    def cantidad_estado(nombre): return int(estados.get(nombre,0))
    indicadores=[
        ('PAQUETES (DEPTO-PQ)',paquetes,'Paquetes con intervenciones'),
        ('INTERVENCIONES TOTALES',total,'Incluye atendidas en otros proyectos'),
        ('CONTRATADOS',cantidad_estado('CONTRATADO'),'Procesos adjudicados'),
        ('POR CONTRATAR',cantidad_estado('EN CONTRATACION'),'En proceso de selección'),
        ('EN EJECUCIÓN',cantidad_estado('EN EJECUCION'),'Obras/actividades activas'),
        ('CULMINADAS',cantidad_estado('CULMINADO'),'Acumulado al corte'),
        ('ATENDIDOS EN OTROS PROYECTOS',cantidad_estado('ATENDIDO EN OTROS PROYECTOS'),'Sin intervención en este proyecto'),
        ('BENEFICIARIOS',f'{u.BENEFICIARIOS.sum():,.0f}','Alcance total'),
        ('HECTÁREAS',f'{u.HECTAREAS.sum():,.2f}','Alcance total'),
        ('KILÓMETROS',f'{u.KM.sum():,.2f}','Alcance total'),
    ]
    for inicio in (0,5):
        columnas=st.columns(5)
        for columna,(nombre,valor,ayuda) in zip(columnas,indicadores[inicio:inicio+5]):
            columna.metric(nombre,valor,help=ayuda)
    st.caption(f'UIR registra {total:,} intervenciones en este territorio; '
               f'{int((u.PAQUETE != "SIN PAQUETE").sum()):,} están asociadas a '
               f'{paquetes:,} paquetes distintos. Las intervenciones atendidas en otros proyectos '
               'se incluyen en el total, pero no necesariamente forman parte de un paquete.')

else:
    k=st.columns(3)
    k[0].metric('Departamentos',territorios.DEPARTAMENTO.nunique())
    k[1].metric('Provincias',len(territorios[NIVELES[:2]].drop_duplicates()))
    k[2].metric('Distritos',len(territorios))
    st.markdown('**🌱 UAT · Intervenciones agropecuarias**')
    k=st.columns(3)
    k[0].metric('Kits',t.KIT.nunique())
    k[1].metric('Intervenciones',len(t))
    k[2].metric('Atenciones programadas',f'{t.PROG.sum():,.0f}')

seccion('04 · Explorador territorial interactivo')
st.caption('Mapa de la unidad seleccionada (o de ambas si no hay selección). Haz clic en un punto o barra para navegar por departamento, provincia y distrito.')

# Cada nivel territorial agrupa las ubicaciones de cada unidad, para evitar
# mostrar centenares de distritos en la vista nacional.
# Si arriba ya se filtró por departamento (o provincia/distrito), el mapa arranca
# en ese lugar y muestra directamente el nivel siguiente (p. ej. las provincias de Lima),
# en vez de un solo punto gigante.
fijo=[]
if dep!='Todo el Perú':
    fijo.append(dep)
    if prov!='Todas las provincias':
        fijo.append(prov)
        if dist!='Todos los distritos': fijo.append(dist)
if 'ruta_mapa_v5' not in st.session_state:
    st.session_state.ruta_mapa_v5 = list(fijo)
firma = (dep, prov, dist, vista)
if st.session_state.get('firma_mapa_v5') != firma:
    st.session_state.ruta_mapa_v5 = list(fijo)
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
if enc2.button('⬆ Subir',disabled=len(ruta)<=len(fijo),key='subir_v5',use_container_width=True):
    st.session_state.ruta_mapa_v5=ruta[:-1];st.rerun()
if enc3.button('🏠 Inicio',disabled=len(ruta)<=len(fijo),key='peru_v5',use_container_width=True):
    st.session_state.ruta_mapa_v5=list(fijo);st.rerun()

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

n_lugares=int(mapa_data.LUGAR.nunique())
altura=int(np.clip(150+42*min(n_lugares,28),340,570))  # pocas barras -> gráfico más bajo y ordenado
izq,der=st.columns([1,1],gap='medium')
with izq:
    with st.container(border=True):
        st.markdown(f'**Cobertura territorial · {vista}**')
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
                map_style='carto-positron',height=altura,
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
                custom_data=['LUGAR'],height=altura,text_auto=',.0f',
                labels={'VALOR':'Beneficiarios / atenciones','ETIQUETA':''})
            estilizar(graf)
            graf.update_layout(margin=dict(l=8,r=12,t=20,b=10),legend=dict(orientation='h',y=1.06))
            graf.update_xaxes(rangemode='tozero');graf.update_layout(bargap=.35)
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

leyenda_unidad = ('<span><b style="color:#1976D2">●</b> UIR · Descolmatación</span>'
                  if vista == 'UIR' else
                  '<span><b style="color:#16A34A">●</b> UAT · Intervenciones</span>')
st.markdown('<div style="display:flex;flex-wrap:wrap;gap:6px 20px;align-items:center;font-size:13px;color:#52685C;margin:8px 2px 14px">'
            + leyenda_unidad +
            '<span style="margin-left:auto;font-size:12px">Clic en un punto o barra para profundizar · ubicaciones sin coordenada exacta son aproximadas</span></div>',
            unsafe_allow_html=True)

if len(ruta)>len(fijo) or len(ruta)==3:
    detalle = por_ruta(u if vista == 'UIR' else t)
    with st.container(border=True):
        st.markdown('**📍 Detalle de '+' › '.join(x.title() for x in ruta)+' · '+vista+'**')
        if detalle.empty:
            st.info('No hay intervenciones de '+vista+' en este territorio.')
        elif vista == 'UIR':
            st.markdown('🌊 **UIR · Descolmatación**')
            m=st.columns(5)
            m[0].metric('Intervenciones',len(detalle))
            m[1].metric('Paquetes',detalle.loc[detalle.PAQUETE.ne('SIN PAQUETE'),'PAQUETE'].nunique())
            m[2].metric('Beneficiarios',f'{detalle.BENEFICIARIOS.sum():,.0f}')
            m[3].metric('Hectáreas',f'{detalle.HECTAREAS.sum():,.2f}')
            m[4].metric('Kilómetros',f'{detalle.KM.sum():,.2f}')
            paquetes_local=detalle[detalle.PAQUETE.ne('SIN PAQUETE')]
            if not paquetes_local.empty:
                resumen_pq=(paquetes_local.groupby('PAQUETE',as_index=False)
                    .agg(Intervenciones=('PAQUETE','size'),Beneficiarios=('BENEFICIARIOS','sum'),
                         Hectáreas=('HECTAREAS','sum'),Kilómetros=('KM','sum')))
                with st.expander('📦 Ver desglose de intervenciones por paquete', expanded=False):
                    st.dataframe(resumen_pq,hide_index=True,use_container_width=True)
            else:
                st.caption('No se identificaron paquetes en este territorio.')
            st.caption(' · '.join(f'{k.title()}: {v}' for k,v in detalle.ESTADO.value_counts().items()))
            extra=['DISTRITO'] if len(ruta)<3 else []
            with st.expander('🌊 Ver tabla UIR (sectores)',expanded=False):
                st.dataframe(detalle[extra+['SECTOR','PAQUETE','ESTADO','BENEFICIARIOS','HECTAREAS','KM']],
                             hide_index=True,use_container_width=True)
        else:
            st.markdown('🌱 **UAT · Intervenciones agropecuarias**')
            m=st.columns(4)
            m[0].metric('Tipos de intervención',detalle.KIT.nunique())
            m[1].metric('Atenciones programadas',f'{detalle.PROG.sum():,.0f}')
            m[2].metric('Distritos',len(detalle[NIVELES].drop_duplicates()))
            m[3].metric('Cumplimiento',f'{detalle.AVANCE.sum()/detalle.PROG.sum()*100:.1f}%' if HAY_AV and detalle.PROG.sum()>0 else 'Sin avance reportado')
            extra=['DISTRITO'] if len(ruta)<3 else []
            cols_t=extra+['KIT','META','PROG']+(['AV_META','AVANCE','PCT'] if HAY_AV else [])
            with st.expander('🌱 Ver tabla UAT (kits)',expanded=False):
                st.dataframe(detalle[cols_t],hide_index=True,use_container_width=True,
                             column_config={'PCT':st.column_config.ProgressColumn('Cumplimiento',min_value=0,max_value=100,format='%.0f%%')})

# La navegación del mapa también filtra TODOS los gráficos y tablas posteriores.
# Se conservan los datos completos en uir/uat para volver a navegar o subir de nivel.
if ruta:
    u=por_ruta(u)
    t=por_ruta(t)
    st.caption('📍 Los análisis siguientes están filtrados por: ' + ' › '.join(x.title() for x in ruta))

sel=st.session_state.v6_sel
dt=t  # valor por defecto; si se abre UAT, se redefine según la intervención elegida

if sel=='UIR':
    PALETA=["#2E7D32","#1565C0","#EF6C00","#6A1B9A","#00838F","#AD1457","#8D6E63","#558B2F","#5C6BC0","#C62828","#00695C","#9E9D24"]
    CMAP_ESTADO={**ESTADOS}
    for i,e in enumerate(sorted(set(u.ESTADO)-set(ESTADOS))): CMAP_ESTADO[e]=PALETA[(i+3)%len(PALETA)]
    CMAP_DEPTO={d:PALETA[i%len(PALETA)] for i,d in enumerate(sorted(u.DEPARTAMENTO.unique()))}
    AV_REAL,AV_PROG='Avance_Fisico_%','Avance_Prog_%'
    MONTOS=['Monto_Contratado','Monto_Ejecutado']
    UMBRAL_ATRASO_PP=10;UMBRAL_DIAS_ESTANCADO=30

    def clave_pq(p):
        m=re.search(r'\d+',p)
        return (re.sub(r'\d+','',p).strip(),int(m.group()) if m else 0,p)
    def tiene(*cols): return all(c in u.columns and u[c].notna().any() for c in cols)
    def pie_fen(data,names,values,title,key,cmap=None,fmt='%{value:,.0f}',height=340):
        f=px.pie(data,names=names,values=values,hole=.45,color=names,color_discrete_map=cmap or {},
                 color_discrete_sequence=PALETA,title=title)
        f.update_traces(texttemplate=fmt+'<br>(%{percent})',textposition='inside',insidetextorientation='horizontal',
                        textfont=dict(color='white',size=12),marker=dict(line=dict(color='white',width=2)),
                        hovertemplate='<b>%{label}</b><br>'+fmt+'<br>%{percent}<extra></extra>',sort=True)
        f.update_layout(height=height,legend=dict(orientation='v',yanchor='middle',y=.5,xanchor='left',x=1.02),
                        uniformtext_minsize=10,uniformtext_mode='hide')
        graficar(f,key)

    es_pq=u.PAQUETE.ne('SIN PAQUETE')
    u_pq=u[es_pq]
    orden_pkg=sorted(u_pq.PAQUETE.unique(),key=clave_pq)

    st.markdown('#### 🌊 UIR · Descolmatación')

    # ───────── ALERTAS: QUÉ DICE EL DASHBOARD ─────────
    # (los estilos .alert-red / .alert-amber / .alert-green están en el bloque TEMA VISUAL)
    seccion('📈 Estado del portafolio · UIR')
    RE_PROBLEMA='ATENDIDO EN OTROS PROYECTOS|PENDIENTE|RIESGO|RETRAS|OBSERV'
    total_u=len(u)
    atencion_u=int(u.ESTADO.str.contains(RE_PROBLEMA,regex=True,na=False).sum())
    ejecucion_u=int(u.ESTADO.str.contains('EJECUCION',na=False).sum())
    pct_atencion=atencion_u/total_u*100 if total_u else 0
    pct_ejec=ejecucion_u/total_u*100 if total_u else 0
    a1,a2,a3=st.columns(3)
    with a1:
        st.markdown(f"""<div class="alert-red"><b>REQUIEREN SEGUIMIENTO</b><br><br>
        {atencion_u:,} registros ({pct_atencion:.1f}%) están atendidos en otros proyectos, pendientes
        o con riesgo/retraso.</div>""",unsafe_allow_html=True)
    with a2:
        cuello=u_pq[u_pq.ESTADO.str.contains('EN CONTRATACION|PENDIENTE|RIESGO|RETRAS|OBSERV',regex=True,na=False)]
        if len(cuello):
            top=cuello.PAQUETE.value_counts()
            msg=(f"<b>{top.index[0]}</b> concentra {top.iloc[0]:,} intervenciones aún sin contratar "
                 f"o con alerta: revisarlo primero.")
        else:
            msg='No hay paquetes con intervenciones por contratar o con alerta con los filtros actuales.'
        st.markdown(f'<div class="alert-amber"><b>PRIORIDAD PMO</b><br><br>{msg}</div>',unsafe_allow_html=True)
    with a3:
        st.markdown(f"""<div class="alert-green"><b>AVANCE</b><br><br>
        {ejecucion_u:,} registros ({pct_ejec:.1f}%) se encuentran en ejecución.</div>""",unsafe_allow_html=True)

    # ───────── DISTRIBUCIÓN TERRITORIAL ─────────
    seccion('🗺️ Distribución territorial · UIR')
    NOMBRES_IND={'Intervenciones':'Intervenciones','BENEFICIARIOS':'Beneficiarios','HECTAREAS':'Hectáreas','KM':'Km'}
    s1,s2=st.columns(2)
    indicador=s1.radio('Indicador a evaluar',list(NOMBRES_IND),format_func=NOMBRES_IND.get,horizontal=True,key='v5_ind')
    deptos_disp=sorted(u.DEPARTAMENTO.dropna().unique())
    if st.session_state.get('v5_dept') not in deptos_disp:
        st.session_state.pop('v5_dept', None)
    dept=s2.selectbox('Departamento para el detalle provincial',deptos_disp,key='v5_dept') if deptos_disp else None
    fmt_ind='%{value:,.0f}' if indicador in ('Intervenciones','BENEFICIARIOS') else '%{value:,.1f}'
    nom=NOMBRES_IND[indicador]
    def serie_por(frame,col):
        s=frame.groupby(col).size().rename('Valor') if indicador=='Intervenciones' else frame.groupby(col)[indicador].sum().rename('Valor')
        return s[s>0].reset_index()
    t1,t2=st.columns(2)
    with t1:
        d=serie_por(u,'DEPARTAMENTO')
        if d.empty: st.warning(f'Sin datos de {nom} para los filtros actuales.')
        else: pie_fen(d,'DEPARTAMENTO','Valor',f'{nom} por departamento','v5_pie_dep',cmap=CMAP_DEPTO,fmt=fmt_ind)
    with t2:
        if dept:
            p=serie_por(u[u.DEPARTAMENTO==dept],'PROVINCIA')
            if p.empty: st.warning(f'Sin datos de {nom} en {dept}.')
            else: pie_fen(p,'PROVINCIA','Valor',f'{nom} por provincia · {dept}','v5_pie_prov',fmt=fmt_ind)
    if dept:
        df_dept=u[u.DEPARTAMENTO==dept]
        u1,u2=st.columns(2)
        with u1:
            est_d=df_dept.groupby('ESTADO').size().reset_index(name='Registros')
            pie_fen(est_d,'ESTADO','Registros',f'Estados de las intervenciones · {dept}','v5_pie_estdep',cmap=CMAP_ESTADO)
        with u2:
            est_p=df_dept.groupby(['PROVINCIA','ESTADO']).size().reset_index(name='Registros')
            fig=px.bar(est_p,y='PROVINCIA',x='Registros',color='ESTADO',orientation='h',barmode='stack',
                       color_discrete_map=CMAP_ESTADO,title=f'Estados por provincia · {dept}')
            fig.update_layout(height=340,xaxis_title=None,yaxis_title=None,legend=dict(orientation='h',y=-.2,title=None))
            graficar(fig,'v5_bar_estprov')

    # ───────── DETALLE POR PAQUETE ─────────
    seccion('🔍 Detalle por paquete · UIR')
    # Aplicar TAMBIÉN el departamento elegido en Distribución territorial.
    # 'u' ya respeta filtros superiores y la navegación del mapa.
    u_pq = u_pq[u_pq.DEPARTAMENTO.eq(dept)].copy() if dept else u_pq.iloc[:0].copy()
    orden_pkg = sorted(u_pq.PAQUETE.unique(), key=clave_pq)
    # Sincronizar el selector con el filtro territorial y la ruta del mapa.
    # Una selección de otro departamento nunca debe mantenerse.
    firma_paquetes = (vista, dep, prov, dist, dept, tuple(ruta), tuple(orden_pkg))
    if st.session_state.get('v9_firma_paquetes') != firma_paquetes:
        st.session_state.pop('v5_pkg', None)
        st.session_state['v9_firma_paquetes'] = firma_paquetes
    if st.session_state.get('v5_pkg') not in orden_pkg:
        st.session_state.pop('v5_pkg', None)
    paquete_detalle=st.selectbox('Selecciona un paquete',orden_pkg,key='v5_pkg') if orden_pkg else None
    if not orden_pkg:
        st.info('No hay paquetes asociados al territorio seleccionado.')
    if paquete_detalle:
        det=u_pq[u_pq.PAQUETE==paquete_detalle]
        q=st.columns(4)
        q[0].metric('Intervenciones',f'{len(det):,}')
        q[1].metric('Beneficiarios',f'{det.BENEFICIARIOS.sum():,.0f}')
        q[2].metric('Hectáreas',f'{det.HECTAREAS.sum():,.2f}')
        q[3].metric('Kilómetros',f'{det.KM.sum():,.2f}')
        v1,v2=st.columns(2)
        with v1:
            d_est=det.groupby('ESTADO').size().reset_index(name='Registros')
            pie_fen(d_est,'ESTADO','Registros',f'Estados · {paquete_detalle}','v5_pie_pq_est',cmap=CMAP_ESTADO)
        with v2:
            d_dis=det.groupby('DISTRITO').size().reset_index(name='Registros')
            pie_fen(d_dis,'DISTRITO','Registros',f'Distritos · {paquete_detalle}','v5_pie_pq_dis')
        with st.expander(f'Ver intervenciones de {paquete_detalle}'):
            cols_det=['PROVINCIA','DISTRITO','SECTOR','KM','BENEFICIARIOS','HECTAREAS','ESTADO']
            st.dataframe(det[cols_det],hide_index=True,use_container_width=True)
            st.download_button('Descargar UIR (CSV)',u[NIVELES+['SECTOR','PAQUETE','ESTADO','BENEFICIARIOS','HECTAREAS','KM']].to_csv(index=False).encode('utf-8-sig'),'uir_filtrado.csv','text/csv',key='v5_csv')

    # ───────── BLOQUES OPCIONALES (se activan si el Excel trae las columnas) ─────────
    if not u_pq.empty and (tiene(AV_REAL,AV_PROG) or tiene(*MONTOS) or tiene('Fecha_Cambio_Estado') or tiene('Restriccion')):
        seccion('⏱️ Cronograma, costos y restricciones · UIR')
        if tiene(AV_REAL,AV_PROG):
            av=u_pq.assign(**{c:pd.to_numeric(u_pq[c],errors='coerce') for c in (AV_REAL,AV_PROG)}).groupby('PAQUETE')[[AV_REAL,AV_PROG]].mean().reindex(orden_pkg).reset_index()
            av['Desviacion']=av[AV_REAL]-av[AV_PROG]
            largo=av.melt(id_vars='PAQUETE',value_vars=[AV_REAL,AV_PROG],var_name='Tipo',value_name='Avance %')
            largo['Tipo']=largo['Tipo'].map({AV_REAL:'Real',AV_PROG:'Programado'})
            fig=px.bar(largo,x='PAQUETE',y='Avance %',color='Tipo',barmode='group',color_discrete_map={'Real':'#2E7D32','Programado':'#90A4AE'},title='Avance físico: real vs programado (promedio por paquete)')
            fig.update_layout(height=340,xaxis_title=None,yaxis_title=None,legend=dict(orientation='h',y=-.35,title=None))
            fig.update_xaxes(type='category',categoryorder='array',categoryarray=orden_pkg,tickangle=-45)
            graficar(fig,'v5_av_real_prog')
            atrasados=av[av.Desviacion<-UMBRAL_ATRASO_PP].sort_values('Desviacion')
            if len(atrasados):
                st.error('PAQUETES ATRASADOS (>%d pp): '%UMBRAL_ATRASO_PP+', '.join(f'{r.PAQUETE} ({r.Desviacion:.0f} pp)' for r in atrasados.itertuples()))
        if tiene(*MONTOS):
            fin=u_pq.assign(**{c:pd.to_numeric(u_pq[c],errors='coerce') for c in MONTOS}).groupby('PAQUETE')[MONTOS].sum().reindex(orden_pkg).reset_index()
            fin['Saldo']=fin.Monto_Contratado-fin.Monto_Ejecutado
            fin['% Ejecución']=(fin.Monto_Ejecutado/fin.Monto_Contratado*100).where(fin.Monto_Contratado>0)
            largo=fin.melt(id_vars='PAQUETE',value_vars=MONTOS,var_name='Concepto',value_name='Monto')
            fig=px.bar(largo,x='PAQUETE',y='Monto',color='Concepto',barmode='group',color_discrete_map={'Monto_Contratado':'#1565C0','Monto_Ejecutado':'#2E7D32'},title='Monto contratado vs ejecutado por paquete')
            fig.update_layout(height=340,xaxis_title=None,yaxis_title=None,legend=dict(orientation='h',y=-.35,title=None))
            fig.update_xaxes(type='category',categoryorder='array',categoryarray=orden_pkg,tickangle=-45)
            graficar(fig,'v5_montos')
            with st.expander('Ver tabla financiera por paquete'):
                st.dataframe(fin,hide_index=True,use_container_width=True)
        if tiene('Fecha_Cambio_Estado'):
            estanc=u_pq[u_pq.ESTADO.str.contains('EN CONTRATACION',na=False)].copy()
            estanc['Dias_en_estado']=(pd.Timestamp.today().normalize()-pd.to_datetime(estanc['Fecha_Cambio_Estado'],errors='coerce',dayfirst=True)).dt.days
            estanc=estanc[estanc.Dias_en_estado>UMBRAL_DIAS_ESTANCADO].sort_values('Dias_en_estado',ascending=False)
            st.warning(f'ESTANCADAS EN CONTRATACIÓN: {len(estanc):,} intervenciones llevan más de {UMBRAL_DIAS_ESTANCADO} días en contratación.')
            if len(estanc):
                with st.expander('Ver intervenciones estancadas'):
                    st.dataframe(estanc[['PAQUETE','DISTRITO','SECTOR','Dias_en_estado']],hide_index=True,use_container_width=True)
        if tiene('Restriccion'):
            rs=u_pq['Restriccion'].dropna().astype(str).str.strip()
            rs=rs[rs!=''].value_counts().reset_index();rs.columns=['Restriccion','Intervenciones']
            if len(rs):
                fig=px.bar(rs,x='Restriccion',y='Intervenciones',color_discrete_sequence=['#E65100'],title='Restricciones más frecuentes')
                fig.update_traces(textposition='outside',cliponaxis=False);fig.update_layout(height=340,xaxis_title=None)
                graficar(fig,'v5_restricciones')

    
# ── Detalle UAT: solo visible al seleccionar UAT ──
elif sel=='UAT':
    st.markdown('#### 🌱 UAT · Intervenciones y metas')
    kit=st.selectbox('Intervención',['Todas las intervenciones']+sorted(t.KIT.unique()),key='v2_kit')
    dt=t if kit=='Todas las intervenciones' else t[t.KIT.eq(kit)]
    ks=st.columns(4)
    ks[0].metric('Intervenciones',dt.KIT.nunique())
    ks[1].metric('Metas',len(dt))
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
    # ───────── UAT POR TERRITORIO E INTERVENCIÓN ─────────
    seccion('📍 UAT por territorio e intervención')
    ini_nv=0 if dep=='Todo el Perú' else (1 if prov=='Todas las provincias' else 2)
    nv=st.radio('Ver por',['Departamento','Provincia','Distrito'],index=ini_nv,horizontal=True,key=f'v7_nivel_uat_{ini_nv}')
    colnv=nv.upper()
    dtt=dt.copy()
    dtt['LUGAR']=dtt.DISTRITO.str.title()+' · '+dtt.PROVINCIA.str.title() if colnv=='DISTRITO' else dtt[colnv].str.title()
    por=dtt.groupby('LUGAR',as_index=False).agg(PROG=('PROG','sum'),AVANCE=('AVANCE','sum'),Intervenciones=('KIT','nunique'))
    por['Porcentaje']=np.where(por.PROG>0,por.AVANCE/por.PROG*100,np.nan)
    por=por.sort_values('PROG',ascending=False)
    top=por.head(25).sort_values('PROG')
    alto=int(np.clip(110+30*len(top),280,820))
    if len(por)>25:
        st.caption(f'Se muestran los 25 con más atenciones programadas de {len(por)} {nv.lower()}s; la tabla incluye todos.')
    cA,cB=st.columns(2)
    with cA:
        if HAY_AV:
            top=top.assign(Estado=top.Porcentaje.map(estado_pct))
            fig=px.bar(top,x='Porcentaje',y='LUGAR',orientation='h',color='Estado',text='Porcentaje',
                       color_discrete_map={x[1]:x[2] for x in SEMAFORO},title=f'Cumplimiento por {nv.lower()} · UAT')
            fig.update_traces(texttemplate='%{text:.1f}%',textposition='outside',cliponaxis=False)
            mx=top.Porcentaje.max()
            fig.update_xaxes(range=[0,max(110,mx*1.15 if pd.notna(mx) else 110)],title='Cumplimiento (%)')
            fig.update_layout(legend=dict(orientation='h',y=-.12,title=None))
        else:
            fig=px.bar(top,x='PROG',y='LUGAR',orientation='h',color_discrete_sequence=['#16A34A'],text='PROG',
                       title=f'Atenciones programadas por {nv.lower()} · UAT')
            fig.update_traces(texttemplate='%{text:,.0f}',textposition='outside',cliponaxis=False)
            fig.update_xaxes(title='Atenciones programadas')
        fig.update_yaxes(title=None,categoryorder='array',categoryarray=list(top.LUGAR))
        fig.update_layout(height=alto,bargap=.3)
        graficar(fig,'v7_uat_terr_pct')
    with cB:
        comp=dtt[dtt.LUGAR.isin(top.LUGAR)].groupby(['LUGAR','KIT'],as_index=False).PROG.sum()
        fig=px.bar(comp,x='PROG',y='LUGAR',color='KIT',orientation='h',barmode='stack',
                   title=f'Qué intervenciones hay en cada {nv.lower()}',labels={'PROG':'Atenciones programadas','KIT':'Intervención'})
        fig.update_yaxes(title=None,categoryorder='array',categoryarray=list(top.LUGAR))
        fig.update_layout(height=alto,bargap=.3,legend=dict(orientation='h',y=-.12,title=None))
        graficar(fig,'v7_uat_terr_kits')
    with st.expander(f'Ver tablas por {nv.lower()} (resumen y por intervención)'):
        tr,tk=st.tabs([f'Resumen por {nv.lower()}','Atenciones por intervención'])
        with tr:
            res=por.rename(columns={'LUGAR':nv,'PROG':'Programado','AVANCE':'Avance','Porcentaje':'Cumplimiento'})
            if not HAY_AV: res=res.drop(columns=['Avance','Cumplimiento'])
            cfg={'Cumplimiento':st.column_config.ProgressColumn('Cumplimiento',min_value=0,max_value=100,format='%.0f%%')} if HAY_AV else {}
            st.dataframe(res,hide_index=True,use_container_width=True,column_config=cfg)
        with tk:
            piv=dtt.pivot_table(index='LUGAR',columns='KIT',values='PROG',aggfunc='sum',fill_value=0).round(0).reset_index().rename(columns={'LUGAR':nv})
            st.dataframe(piv,hide_index=True,use_container_width=True)


    with st.expander('Ver tabla de intervenciones UAT'):
        cols_t=NIVELES+['KIT','META','PROG']+(['AV_META','AVANCE','PCT'] if HAY_AV else [])
        st.dataframe(dt[cols_t],hide_index=True,use_container_width=True)


# ═════════════ SECCIONES FINALES ═════════════

with st.expander('Notas metodológicas y calidad de datos'):
    st.markdown('**UIR:** se excluye del impacto el estado «Atendido en otros proyectos». **UAT:** Praderas = meta física ÷ 5 × 3. **Cobertura:** distritos únicos entre unidades. **Beneficiarios:** no se suman UIR y UAT, pues no se ha validado que sean personas únicas ni indicadores equivalentes.')
    if SIN_COINCIDENCIA:st.warning(f'{SIN_COINCIDENCIA} combinaciones de avance UAT no coinciden con la programación.')
st.caption('Fuentes: UIR · EXCEL DS 164 EF | UAT · datos_fen_2026, plantilla de avance y coordenadas distritales.')
