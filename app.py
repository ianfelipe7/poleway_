"""
PoleWay - ANALISADOR DE ROTAS (OSRM)
"""

import pandas as pd
import numpy as np
import streamlit as st
import folium
import requests
from scipy.spatial.distance import euclidean #type: ignore
from streamlit_folium import st_folium       #type: ignore
import os
from datetime import datetime
import polyline                              #type: ignore

# ── Configs
st.set_page_config(
    page_title="PoleWay · Rotas",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_CLIENTES          = os.path.join(_DIR, "rotas_processadas_320.xlsx")
ARQUIVO_TEMPO_ATENDIMENTO = os.path.join(_DIR, "segmento.xlsx")
ARQUIVO_VENDEDORES        = os.path.join(_DIR, "endereco.xlsx")
ARQUIVO_FATURAMENTO       = os.path.join(_DIR, "faturamento.xlsx")
ARQUIVO_PERNOITES         = os.path.join(_DIR, "pernoites.xlsx")
COLUNA_ROTA               = "Rota"
OSRM_BASE_URL             = "https://router.project-osrm.org"

ORDEM_DIA_SEMANA = {1:'Seg', 2:'Ter', 3:'Qua', 4:'Qui', 5:'Sex', 6:'Sáb', 7:'Dom'}
ORDEM_INVERSA    = {v: k for k, v in ORDEM_DIA_SEMANA.items()}

CITY_COLORS = [
    '#e74c3c','#3498db','#2ecc71','#f39c12','#9b59b6',
    '#1abc9c','#e67e22','#e91e63','#00bcd4','#8bc34a',
    '#ff5722','#607d8b','#ffeb3b','#795548','#03a9f4',
]

# ── Design
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

*, *::before, *::after { box-sizing: border-box; }

html, body, .stApp {
    font-family: 'DM Sans', sans-serif !important;
    background: #0f1117 !important;
    color: #e8eaf0 !important;
}

[data-testid="stSidebar"] {
    background: #161b27 !important;
    border-right: 1px solid #1e2535 !important;
}
[data-testid="stSidebar"] * { color: #c8ccd8 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #ffffff !important; }

[data-testid="stSidebar"] .stButton > button {
    background: linear-gradient(135deg, #c0392b, #e55a1f) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.03em !important;
    transition: opacity .2s !important;
}
[data-testid="stSidebar"] .stButton > button:hover { opacity: 0.85 !important; }

[data-testid="stMetric"] {
    background: #161b27 !important;
    border: 1px solid #1e2535 !important;
    border-radius: 8px !important;
    padding: 14px 18px !important;
}
[data-testid="stMetricValue"] { color: #ffffff !important; font-size: 1.35rem !important; font-weight: 600 !important; }
[data-testid="stMetricLabel"] { color: #8892a4 !important; font-size: 0.72rem !important; text-transform: uppercase; letter-spacing: .06em; }

.rota-card {
    background: linear-gradient(100deg, #1a0a0e 0%, #1a140c 100%);
    border: 1px solid #2e1a1a;
    border-left: 3px solid #c0392b;
    border-radius: 8px;
    padding: 13px 20px;
    margin: 22px 0 8px 0;
    display: flex;
    align-items: center;
    gap: 14px;
}
.rota-card-num  { font-family: 'DM Mono', monospace; font-size: 0.68rem; color: #c0392b; text-transform: uppercase; letter-spacing: .14em; }
.rota-card-title{ font-size: .98rem; font-weight: 600; color: #f0f2f8; margin-top: 2px; }
.rota-card-meta { font-size: .78rem; color: #8892a4; margin-left: auto; }

[data-testid="stDataFrame"] { border-radius: 8px !important; overflow: hidden !important; }

.legenda-row { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 8px; align-items: center; }
.legenda-dot { display: inline-flex; align-items: center; gap: 5px; font-size: .74rem; color: #8892a4; }
.legenda-dot span { width: 9px; height: 9px; border-radius: 50%; display: inline-block; flex-shrink: 0; }

.page-title { font-size: 1.5rem; font-weight: 600; color: #f0f2f8; margin-bottom: 1px; }
.page-sub   { font-size: .8rem; color: #8892a4; margin-bottom: 18px; }

.stDownloadButton > button {
    background: #1e2535 !important;
    color: #c8ccd8 !important;
    border: 1px solid #2a3347 !important;
    border-radius: 6px !important;
    font-size: .82rem !important;
    transition: border-color .2s !important;
}
.stDownloadButton > button:hover { border-color: #c0392b !important; color: #fff !important; }

hr { border-color: #1e2535 !important; margin: 18px 0 !important; }
[data-testid="stProgressBar"] > div > div { background: #c0392b !important; }
</style>
""", unsafe_allow_html=True)

# ── API 
def testar_osrm_api():
    try:
        r = requests.get(
            f"{OSRM_BASE_URL}/route/v1/driving/-39.3153,-7.2131;-39.3253,-7.2231",
            params={'overview': 'false'}, timeout=5
        )
        return r.status_code == 200 and r.json().get('code') == 'Ok'
    except:
        return False

def calcular_distancia_osrm(origem, destino):
    try:
        url = f"{OSRM_BASE_URL}/route/v1/driving/{origem[1]},{origem[0]};{destino[1]},{destino[0]}"
        r = requests.get(url, params={'overview':'false','steps':'false'}, timeout=10)
        if r.status_code == 200:
            data = r.json()
            if data.get('code') == 'Ok' and data.get('routes'):
                rt = data['routes'][0]
                return rt['distance']/1000, rt['duration']/60
    except:
        pass
    return calcular_distancia_euclidiana(origem, destino)

def calcular_distancia_euclidiana(origem, destino):
    try:
        d = euclidean(origem, destino) * 111 * 1.3
        return d, (d/50)*60
    except:
        return 0, 0

def obter_rota_osrm(origem, destino):
    try:
        url = f"{OSRM_BASE_URL}/route/v1/driving/{origem[1]},{origem[0]};{destino[1]},{destino[0]}"
        r = requests.get(url, params={'overview':'full','geometries':'polyline'}, timeout=10)
        if r.status_code == 200:
            data = r.json()
            if data.get('code') == 'Ok' and data.get('routes'):
                geo = data['routes'][0].get('geometry')
                if geo:
                    return [[c[0], c[1]] for c in polyline.decode(geo)]
    except:
        pass
    return None

# ── Funções auxiliares 
def normalizar_cidade(nome):
    if pd.isna(nome):
        return 'DESCONHECIDA'
    nome = str(nome).upper().strip()
    for o, n in {'Á':'A','À':'A','Ã':'A','Â':'A','É':'E','Ê':'E','Í':'I',
                 'Ó':'O','Õ':'O','Ô':'O','Ú':'U','Ç':'C'}.items():
        nome = nome.replace(o, n)
    return nome

def detectar_coluna(df, nomes):
    for n in nomes:
        if n in df.columns:
            return n
    for col in df.columns:
        for n in nomes:
            if n.lower() in col.lower():
                return col
    return None

def formatar_moeda(valor):
    if pd.isna(valor) or valor == 0:
        return "R$ 0"
    return f"R$ {valor:,.0f}".replace(',','X').replace('.',',').replace('X','.')

def minutos_para_hhmm(m):
    if pd.isna(m) or m == 0:
        return "00:00"
    return f"{int(m//60):02d}:{int(m%60):02d}"

def order_nearest_neighbor(clientes, start):
    if not clientes:
        return []
    unvisited, ordered, current = clientes.copy(), [], start
    while unvisited:
        nearest = min(unvisited, key=lambda c: euclidean(current, (c['latitude'], c['longitude'])))
        ordered.append(nearest)
        unvisited.remove(nearest)
        current = (nearest['latitude'], nearest['longitude'])
    return ordered

# ── Carga de dados ─────────────────────────────────────────────────────
@st.cache_data
def _load_pernoites_raw():
    if not os.path.exists(ARQUIVO_PERNOITES):
        return {}, None
    try:
        df = pd.read_excel(ARQUIVO_PERNOITES, engine='openpyxl')
        df.columns = [c.strip().lower() for c in df.columns]
        df['rota'] = pd.to_numeric(df['rota'], errors='coerce')
        df['dia']  = pd.to_numeric(df['dia'],  errors='coerce')
        df['latitude']  = pd.to_numeric(
            df['latitude'].astype(str).str.replace('"','',regex=False).str.strip(), errors='coerce')
        df['longitude'] = pd.to_numeric(
            df['longitude'].astype(str).str.replace('"','',regex=False).str.strip(), errors='coerce')
        df = df.dropna(subset=['rota','dia','latitude','longitude'])

        resultado = {}
        for rota_num, grp in df.groupby('rota'):
            grp = grp.sort_values('dia')
            rota_info = {}
            ini_row = grp[grp['ponto'].str.lower() == 'inicio']
            inicio_coords = None
            if not ini_row.empty:
                r = ini_row.iloc[0]
                desc = str(r.get('cidade','')) + (f" - {r.get('hotel','')}" if r.get('hotel') else '')
                inicio_coords = (float(r['latitude']), float(r['longitude']), desc.strip(' -'))

            pern = grp[grp['ponto'].str.lower() == 'pernoite'].sort_values('dia')
            dias = sorted(pern['dia'].unique())
            for i, dia in enumerate(dias):
                p = pern[pern['dia'] == dia].iloc[0]
                desc_p = str(p.get('cidade','')) + (f" - {p.get('hotel','')}" if p.get('hotel') else '')
                termino = (float(p['latitude']), float(p['longitude']), desc_p.strip(' -'))
                ini = inicio_coords if i == 0 else rota_info[dias[i-1]]['termino']
                rota_info[dia] = {'inicio': ini if ini else termino, 'termino': termino}
            resultado[rota_num] = rota_info
        return resultado, None
    except Exception as e:
        return {}, str(e)

def load_pernoites():
    dados, erro = _load_pernoites_raw()
    if erro:
        st.error(f"Erro ao carregar pernoites: {erro}")
    return dados

@st.cache_data
def load_faturamento():
    if not os.path.exists(ARQUIVO_FATURAMENTO):
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_r = pd.read_excel(ARQUIVO_FATURAMENTO, sheet_name='Rotas',   engine='openpyxl')
        df_c = pd.read_excel(ARQUIVO_FATURAMENTO, sheet_name='Cidades', engine='openpyxl')
        for df in [df_r, df_c]:
            if 'Rota'        in df.columns: df['Rota']        = pd.to_numeric(df['Rota'],        errors='coerce')
            if 'Faturamento' in df.columns: df['Faturamento'] = pd.to_numeric(df['Faturamento'], errors='coerce')
        if 'Cidade' in df_c.columns:
            df_c['Cidade'] = df_c['Cidade'].apply(normalizar_cidade)
        for col in ['Dia','dia','DIA','Ordem','ordem']:
            if col in df_c.columns:
                df_c['DiaNum'] = pd.to_numeric(df_c[col], errors='coerce')
                break
        for col in ['cdCliente','Cliente','cliente','Codigo','codigo']:
            if col in df_c.columns:
                df_c['CdCliente'] = df_c[col].astype(str)
                break
        return df_r, df_c
    except:
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data
def load_vendedores():
    if not os.path.exists(ARQUIVO_VENDEDORES):
        return pd.DataFrame()
    try:
        df = pd.read_excel(ARQUIVO_VENDEDORES, engine='openpyxl')
        col_lat  = detectar_coluna(df, ['latitude','Latitude','lat'])
        col_lon  = detectar_coluna(df, ['longitude','Longitude','lon','lng'])
        col_rota = detectar_coluna(df, ['rota','Rota','ROTA'])
        if not col_lat or not col_lon:
            return pd.DataFrame()
        df = df.rename(columns={col_lat:'latitude', col_lon:'longitude'})
        if col_rota:
            df = df.rename(columns={col_rota:'Rota'})
        return df[df['latitude'].notna() & df['longitude'].notna()].copy()
    except:
        return pd.DataFrame()

@st.cache_data
def load_clientes():
    if not os.path.exists(ARQUIVO_CLIENTES):
        return pd.DataFrame()
    try:
        df = pd.read_excel(ARQUIVO_CLIENTES, engine='openpyxl')
        if 'latitude' not in df.columns or 'longitude' not in df.columns:
            return pd.DataFrame()
        df = df[df['latitude'].notna() & df['longitude'].notna()].copy()
        col_rota = detectar_coluna(df, [COLUNA_ROTA,'rota','Rota','ROTA'])
        if not col_rota:
            return pd.DataFrame()
        df['ROTA'] = df[col_rota]
        col_ordem = detectar_coluna(df, ['Ordem','ordem','ORDEM'])
        if col_ordem:
            df['DiaSemana'] = pd.to_numeric(df[col_ordem], errors='coerce').map(ORDEM_DIA_SEMANA).fillna('INDEFINIDO')
            df['DiaNum']    = pd.to_numeric(df[col_ordem], errors='coerce')
        else:
            df['DiaSemana'] = 'INDEFINIDO'
            df['DiaNum']    = np.nan
        col_cidade = detectar_coluna(df, ['cidade','Cidade','dsCidadeComercial','nmCidade','municipio'])
        df['Cidade'] = df[col_cidade].apply(normalizar_cidade) if col_cidade else 'DESCONHECIDA'
        return _carregar_tempo_atendimento(df)
    except:
        return pd.DataFrame()

def _carregar_tempo_atendimento(df):
    if not os.path.exists(ARQUIVO_TEMPO_ATENDIMENTO):
        df['TempoVisita'] = 30
        return df
    try:
        dt = pd.read_excel(ARQUIVO_TEMPO_ATENDIMENTO, engine='openpyxl')
        col_seg  = next((c for c in dt.columns if 'segmento' in c.lower()), None)
        col_tmp  = next((c for c in dt.columns if 'tempo' in c.lower() and 'visita' in c.lower()), None)
        col_scli = next((c for c in df.columns  if 'segmento' in c.lower()), None)
        if col_seg and col_tmp and col_scli:
            mapa = dict(zip(
                dt[col_seg].astype(str).str.strip().str.upper(),
                pd.to_numeric(dt[col_tmp], errors='coerce').fillna(30)
            ))
            df['TempoVisita'] = df[col_scli].astype(str).str.strip().str.upper().map(mapa).fillna(30)
        else:
            df['TempoVisita'] = 30
    except:
        df['TempoVisita'] = 30
    return df

# ── Análise de rota
def analisar_rota(df_rota, rota_num, vendor_home, df_fat_cidades, pernoites_data, usar_osrm=True):
    if df_rota.empty:
        return None, None

    dias_da_rota = sorted(
        df_rota['DiaSemana'].dropna().unique().tolist(),
        key=lambda d: ORDEM_INVERSA.get(d, 99)
    )
    col_cli  = detectar_coluna(df_rota, ['cdCliente','codigo','Codigo','Cliente','id'])
    col_nome = detectar_coluna(df_rota, ['nmFantasia','nome','Nome','razao_social'])
    pernoites_rota = pernoites_data.get(rota_num, {})
    calcular = calcular_distancia_osrm if usar_osrm else calcular_distancia_euclidiana

    resultados, todos_clientes = [], []

    for idx, dia in enumerate(dias_da_rota):
        dia_num = ORDEM_INVERSA.get(dia)
        df_dia  = df_rota[df_rota['DiaSemana'] == dia].copy()
        if df_dia.empty:
            continue

        cor = CITY_COLORS[idx % len(CITY_COLORS)]

        # Início e término
        info_dia       = pernoites_rota.get(dia_num, {})
        inicio_coords  = info_dia.get('inicio',  (vendor_home[0], vendor_home[1], 'Base'))
        termino_coords = info_dia.get('termino', (vendor_home[0], vendor_home[1], 'Base'))
        inicio_latlon  = (inicio_coords[0],  inicio_coords[1])
        termino_latlon = (termino_coords[0], termino_coords[1])
        inicio_desc    = inicio_coords[2]  if len(inicio_coords)  > 2 else ''
        termino_desc   = termino_coords[2] if len(termino_coords) > 2 else ''

        clientes_ord = order_nearest_neighbor(df_dia.to_dict('records'), inicio_latlon)

        km_total = tempo_atend = tempo_desloc = 0.0

        # Faturamento cliente/dia
        faturamento_dia = 0
        if not df_fat_cidades.empty:
            tem_cli = 'CdCliente' in df_fat_cidades.columns and col_cli
            tem_dia = 'DiaNum'    in df_fat_cidades.columns
            if tem_cli and tem_dia:
                cods = {str(c.get(col_cli,'')) for c in clientes_ord}
                mask = (
                    (df_fat_cidades['Rota']      == rota_num) &
                    (df_fat_cidades['DiaNum']    == dia_num) &
                    (df_fat_cidades['CdCliente'].isin(cods))
                )
            else:
                cidades_dia = df_dia['Cidade'].dropna().unique().tolist()
                mask = (
                    (df_fat_cidades['Rota']   == rota_num) &
                    (df_fat_cidades['Cidade'].isin(cidades_dia))
                )
            fat_f = df_fat_cidades[mask]
            if not fat_f.empty:
                faturamento_dia = fat_f['Faturamento'].sum()

        for i, cli in enumerate(clientes_ord):
            cli.update({
                'cor': cor, 'cidade_display': dia, 'ordem': i+1,
                'cod_cliente':  cli.get(col_cli,  '') if col_cli  else '',
                'nome_cliente': cli.get(col_nome, '') if col_nome else '',
                'inicio_dia':   inicio_coords,
                'termino_dia':  termino_coords,
            })
            tv = cli.get('TempoVisita', 30)
            tempo_atend += float(tv) if not pd.isna(tv) and tv > 0 else 30

        # Fluxo início→clientes→término
        if clientes_ord:
            d, t = calcular(inicio_latlon, (clientes_ord[0]['latitude'], clientes_ord[0]['longitude']))
            km_total += d; tempo_desloc += t
        for i in range(len(clientes_ord)-1):
            d, t = calcular(
                (clientes_ord[i]['latitude'],   clientes_ord[i]['longitude']),
                (clientes_ord[i+1]['latitude'], clientes_ord[i+1]['longitude'])
            )
            km_total += d; tempo_desloc += t
        if clientes_ord:
            d, t = calcular((clientes_ord[-1]['latitude'], clientes_ord[-1]['longitude']), termino_latlon)
            km_total += d; tempo_desloc += t

        resultados.append({
            'Dia': dia, 'Início': inicio_desc, 'Término': termino_desc,
            'Clientes': len(clientes_ord), 'KM': round(km_total, 1),
            'Atend.': int(round(tempo_atend)), 'Desloc.': int(round(tempo_desloc)),
            'Total': int(round(tempo_atend + tempo_desloc)),
            'Faturamento': faturamento_dia, 'cor': cor,
        })
        todos_clientes.extend(clientes_ord)

    return pd.DataFrame(resultados), todos_clientes

# ── Mapa
def criar_mapa(clientes, vendor_home, usar_osrm=True):
    if not clientes:
        return folium.Map(location=[-7.2, -39.3], zoom_start=9, tiles='CartoDB positron')

    lats = [c['latitude'] for c in clientes]
    lons = [c['longitude'] for c in clientes]
    m = folium.Map(location=[np.mean(lats), np.mean(lons)], zoom_start=10, tiles='CartoDB positron')

    por_dia = {}
    for c in clientes:
        por_dia.setdefault(c.get('cidade_display','?'), []).append(c)

    marcados_ini, marcados_ter = set(), set()

    for dia, lista in por_dia.items():
        lista = sorted(lista, key=lambda x: x.get('ordem', 0))
        cor   = lista[0].get('cor', '#999')
        ini   = lista[0].get('inicio_dia')
        ter   = lista[0].get('termino_dia')

        def add_linha(a, b, dash=None):
            coords = obter_rota_osrm(a, b) if usar_osrm else None
            if coords and len(coords) > 1:
                folium.PolyLine(coords, color=cor, weight=4 if dash else 5,
                                opacity=0.85, dash_array=dash).add_to(m)
            else:
                folium.PolyLine([list(a), list(b)], color=cor,
                                weight=3, opacity=0.5, dash_array='6 8').add_to(m)

        if ini:
            add_linha((ini[0], ini[1]), (lista[0]['latitude'], lista[0]['longitude']), dash='8 5')
        for i in range(len(lista)-1):
            add_linha((lista[i]['latitude'],   lista[i]['longitude']),
                      (lista[i+1]['latitude'], lista[i+1]['longitude']))
        if ter:
            add_linha((lista[-1]['latitude'], lista[-1]['longitude']), (ter[0], ter[1]), dash='8 5')

        if ini and dia not in marcados_ini:
            folium.Marker([ini[0], ini[1]],
                popup=folium.Popup(f"<b>Início · {dia}</b><br>{ini[2] if len(ini)>2 else ''}", max_width=220),
                tooltip=f"▶ Início {dia}",
                icon=folium.Icon(color='green', icon='play', prefix='fa')
            ).add_to(m)
            marcados_ini.add(dia)

        for cli in lista:
            folium.CircleMarker(
                [cli['latitude'], cli['longitude']],
                radius=7, color=cor, fill=True, fill_color=cor, fill_opacity=0.9,
                popup=folium.Popup(
                    f"<div style='font-family:sans-serif;min-width:170px'>"
                    f"<b>{cli.get('cod_cliente','')}</b><br>"
                    f"{cli.get('nome_cliente','')}<br>"
                    f"<small style='color:#888'>{dia} · visita {cli.get('ordem','')}</small></div>",
                    max_width=250
                ),
                tooltip=f"{cli.get('ordem','')}. {cli.get('cod_cliente','')} — {dia}"
            ).add_to(m)

        if ter and dia not in marcados_ter:
            folium.Marker([ter[0], ter[1]],
                popup=folium.Popup(f"<b>Término · {dia}</b><br>{ter[2] if len(ter)>2 else ''}", max_width=220),
                tooltip=f"⏹ Término {dia}",
                icon=folium.Icon(color='red', icon='stop', prefix='fa')
            ).add_to(m)
            marcados_ter.add(dia)

    return m

# ── Relatório HTML
def gerar_html_relatorio(dados_rotas, incluir_mapas=True):
    data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")
    n = len(dados_rotas)
    med_cli = sum(r['total_clientes'] for r in dados_rotas) / n if n else 0
    med_km  = sum(r['total_km']       for r in dados_rotas) / n if n else 0
    med_tmp = sum(r['total_tempo']    for r in dados_rotas) / n if n else 0
    med_fat = sum(r.get('fat_rota',0) for r in dados_rotas) / n if n else 0

    rows_html = ''
    for rota in dados_rotas:
        mapa_html = ''
        if incluir_mapas and rota.get('mapa'):
            try:
                mapa_html = f"<div class='mapa-wrap'>{rota['mapa']._repr_html_()}</div>"
            except:
                pass

        linhas = ''
        for _, row in rota['tabela'].iterrows():
            is_tot = row.get('Dia','') == 'TOTAL'
            cls    = ' class="tot"' if is_tot else ''
            linhas += (
                f"<tr{cls}><td>{row.get('Dia','')}</td>"
                f"<td>{row.get('Início','')}</td><td>{row.get('Término','')}</td>"
                f"<td>{row['Clientes']}</td><td>{row['KM']}</td>"
                f"<td>{row['Atend.']}</td><td>{row['Desloc.']}</td>"
                f"<td>{row['Total']}</td><td>{row['Faturamento']}</td></tr>"
            )

        legenda_html = ''.join(
            f'<span class="ldot"><span style="background:{r["cor"]};width:10px;height:10px;'
            f'border-radius:50%;display:inline-block;margin-right:5px"></span>{r["Dia"]}</span>'
            for _, r in rota['resultado'].iterrows() if r['Clientes'] > 0
        )

        rows_html += f"""
        <div class="rota-sec">
          <div class="rota-hdr">Rota {rota['numero']}
            <span class="rota-meta">{rota['vendedor']} · {formatar_moeda(rota.get('fat_rota',0))}</span>
          </div>
          <div class="rota-body">
            <div class="kpis">
              <div class="kpi"><div class="v">{rota['total_clientes']}</div><div class="l">Clientes</div></div>
              <div class="kpi"><div class="v">{rota['total_km']:.1f} km</div><div class="l">Distância</div></div>
              <div class="kpi"><div class="v">{minutos_para_hhmm(rota['total_atend'])}</div><div class="l">Atendimento</div></div>
              <div class="kpi"><div class="v">{minutos_para_hhmm(rota['total_desloc'])}</div><div class="l">Deslocamento</div></div>
              <div class="kpi"><div class="v">{minutos_para_hhmm(rota['total_tempo'])}</div><div class="l">Total</div></div>
              <div class="kpi"><div class="v">{formatar_moeda(rota.get('fat_rota',0))}</div><div class="l">Faturamento</div></div>
            </div>
            <div class="layout">
              {mapa_html}
              <div class="tbl-wrap"><table>
                <thead><tr><th>Dia</th><th>Início</th><th>Término</th>
                  <th>Cli.</th><th>KM</th><th>Atend.</th><th>Desloc.</th><th>Total</th><th>Fat.</th>
                </tr></thead>
                <tbody>{linhas}</tbody>
              </table></div>
            </div>
            <div class="leg">{legenda_html}
              <span class="ldot">🟢 Início &nbsp;🔴 Término</span>
            </div>
          </div>
        </div>"""

    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>PoleWay · Relatório</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',sans-serif;background:#0f1117;color:#e8eaf0;padding:24px}}
.hdr{{background:linear-gradient(120deg,#c0392b,#e55a1f);padding:24px 32px;border-radius:10px;margin-bottom:24px}}
.hdr h1{{font-size:1.5rem;font-weight:700}} .hdr small{{opacity:.8;font-size:.85rem}}
.kpis-g{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:24px}}
.kpi{{background:#161b27;border:1px solid #1e2535;border-radius:8px;padding:14px;text-align:center}}
.kpi .v{{font-size:1.4rem;font-weight:700;color:#fff}}
.kpi .l{{font-size:.7rem;color:#8892a4;text-transform:uppercase;letter-spacing:.06em;margin-top:2px}}
.rota-sec{{background:#161b27;border:1px solid #1e2535;border-radius:10px;margin-bottom:20px;overflow:hidden}}
.rota-hdr{{background:linear-gradient(100deg,#1e0a0a,#1e1208);border-left:3px solid #c0392b;padding:12px 20px;font-weight:600;font-size:.95rem}}
.rota-meta{{float:right;font-weight:400;font-size:.8rem;color:#8892a4}}
.rota-body{{padding:18px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:10px;margin-bottom:16px}}
.layout{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.mapa-wrap{{height:380px;border-radius:8px;overflow:hidden;border:1px solid #1e2535}}
.mapa-wrap iframe{{width:100%;height:100%;border:none}}
.tbl-wrap{{overflow-x:auto}}
.leg{{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-top:14px;padding-top:12px;border-top:1px solid #1e2535;font-size:.76rem;color:#8892a4}}
.ldot{{display:inline-flex;align-items:center;gap:5px}}
table{{width:100%;border-collapse:collapse;font-size:.77rem}}
th{{background:#1e2535;color:#8892a4;padding:8px 10px;text-align:left;font-size:.69rem;text-transform:uppercase;letter-spacing:.05em}}
td{{padding:8px 10px;border-bottom:1px solid #1e2535;color:#c8ccd8}}
tr.tot td{{background:#2a1e0a;color:#f5c842;font-weight:600;border:none}}
tr:hover td{{background:#1a2030}}
@media(max-width:860px){{.layout{{grid-template-columns:1fr}}}}
@media print{{body{{background:#fff;color:#000}}.rota-hdr,.hdr{{-webkit-print-color-adjust:exact;print-color-adjust:exact}}}}
</style></head><body>
<div class="hdr"><h1>🚚 PoleWay · Relatório de Rotas</h1><small>Gerado em {data_geracao} · OSRM</small></div>
<div class="kpis-g">
  <div class="kpi"><div class="v">{n}</div><div class="l">Rotas</div></div>
  <div class="kpi"><div class="v">{med_cli:.0f}</div><div class="l">Média Clientes</div></div>
  <div class="kpi"><div class="v">{med_km:.1f} km</div><div class="l">Distância Média</div></div>
  <div class="kpi"><div class="v">{minutos_para_hhmm(med_tmp)}</div><div class="l">Tempo Médio</div></div>
  <div class="kpi"><div class="v">{formatar_moeda(med_fat)}</div><div class="l">Fat. Médio</div></div>
</div>
{rows_html}
<p style="text-align:center;color:#3d4f6e;font-size:.75rem;margin-top:24px">PoleWay · OSRM · {data_geracao}</p>
</body></html>"""

# ── Interface
def main():
    df_clientes              = load_clientes()
    df_vendedores            = load_vendedores()
    df_fat_rotas, df_fat_cidades = load_faturamento()
    pernoites_data           = load_pernoites()
    osrm_ok                  = testar_osrm_api()

    if df_clientes.empty:
        st.error("Arquivo de clientes não encontrado ou inválido.")
        st.stop()

    rotas_disponiveis = sorted(df_clientes['ROTA'].dropna().unique().tolist())

    # Sidebar
    with st.sidebar:
        st.markdown("## Configurações")
        ca, cb = st.columns(2)
        ca.success("OSRM ✓"      if osrm_ok         else "OSRM ✗")
        cb.success("Pernoites ✓" if pernoites_data   else "Pernoites ✗")
        if not df_fat_rotas.empty:
            st.success(f"Faturamento · {len(df_fat_rotas)} rotas")

        usar_osrm = st.checkbox("Usar OSRM API", value=osrm_ok,
                                help="Desmarque para cálculo aproximado (mais rápido)")
        st.markdown("### Rotas")
        rotas_selecionadas = st.multiselect(
            "Selecionar rotas:",
            options=rotas_disponiveis,
            default=rotas_disponiveis[:3] if len(rotas_disponiveis) > 3 else rotas_disponiveis,
            format_func=lambda x: f"Rota {int(x)}" if pd.notna(x) else "—"
        )
        st.markdown("---")
        analisar = st.button("Analisar Rotas", use_container_width=True, type="primary")

    if 'dados_exportar' not in st.session_state:
        st.session_state.dados_exportar = []

    # Header
    st.markdown('<div class="page-title">🚚 PoleWay · Análise de Rotas</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub">Roteamento via OSRM · Open Source Routing Machine</div>', unsafe_allow_html=True)

    # Análise
    if analisar:
        st.session_state.dados_exportar = []

        med_fat_g = 0
        if not df_fat_rotas.empty:
            fat_sel = df_fat_rotas[df_fat_rotas['Rota'].isin(rotas_selecionadas)]
            med_fat_g = fat_sel['Faturamento'].mean() if not fat_sel.empty else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Clientes",     len(df_clientes))
        c2.metric("Rotas totais", len(rotas_disponiveis))
        c3.metric("Vendedores",   len(df_vendedores))
        c4.metric("Fat. médio",   formatar_moeda(med_fat_g))

        st.markdown("---")
        prog = st.progress(0)
        stat = st.empty()

        for i, rota_num in enumerate(sorted(rotas_selecionadas)):
            prog.progress((i+1) / len(rotas_selecionadas))
            stat.caption(f"Processando rota {int(rota_num)}…")

            df_rota = df_clientes[df_clientes['ROTA'] == rota_num].copy()
            if df_rota.empty:
                continue

            vendor_home   = (-7.2131, -39.3153)
            nome_vendedor = "—"
            if not df_vendedores.empty and 'Rota' in df_vendedores.columns:
                vend = df_vendedores[df_vendedores['Rota'] == rota_num]
                if not vend.empty:
                    vendor_home = (vend.iloc[0]['latitude'], vend.iloc[0]['longitude'])
                    col_nv = detectar_coluna(vend, ['nome','Nome','vendedor','Vendedor'])
                    if col_nv:
                        nome_vendedor = vend.iloc[0][col_nv]

            fat_rota = 0
            if not df_fat_rotas.empty:
                fr = df_fat_rotas[df_fat_rotas['Rota'] == rota_num]
                if not fr.empty:
                    fat_rota = fr['Faturamento'].iloc[0]

            df_resultado, clientes = analisar_rota(
                df_rota, rota_num, vendor_home, df_fat_cidades, pernoites_data, usar_osrm
            )
            if df_resultado is None or df_resultado.empty:
                continue

            total_cli = df_resultado['Clientes'].sum()
            if total_cli == 0:
                continue

            st.markdown(f"""
            <div class="rota-card">
              <div>
                <div class="rota-card-num">Rota {int(rota_num)}</div>
                <div class="rota-card-title">{nome_vendedor} · {total_cli} clientes</div>
              </div>
              <div class="rota-card-meta">{formatar_moeda(fat_rota)}</div>
            </div>""", unsafe_allow_html=True)

            col_mapa, col_tab = st.columns([1, 1])

            with col_mapa:
                mapa = criar_mapa(clientes, vendor_home, usar_osrm)
                st_folium(mapa, width=None, height=420, key=f"map_{rota_num}", returned_objects=[])

            with col_tab:
                df_tab = df_resultado[['Dia','Início','Término','Clientes','KM',
                                       'Atend.','Desloc.','Total','Faturamento']].copy()

                tot_cli = df_tab['Clientes'].sum()
                tot_km  = df_tab['KM'].sum()
                tot_at  = df_tab['Atend.'].sum()
                tot_dl  = df_tab['Desloc.'].sum()
                tot_tt  = df_tab['Total'].sum()
                tot_fat = df_tab['Faturamento'].sum()

                df_tab['KM']          = df_tab['KM'].apply(lambda x: f"{x:.1f}")
                df_tab['Atend.']      = df_tab['Atend.'].apply(minutos_para_hhmm)
                df_tab['Desloc.']     = df_tab['Desloc.'].apply(minutos_para_hhmm)
                df_tab['Total']       = df_tab['Total'].apply(minutos_para_hhmm)
                df_tab['Faturamento'] = df_tab['Faturamento'].apply(formatar_moeda)

                total_row = {
                    'Dia':'TOTAL','Início':'—','Término':'—',
                    'Clientes': tot_cli, 'KM': f"{tot_km:.1f}",
                    'Atend.': minutos_para_hhmm(tot_at), 'Desloc.': minutos_para_hhmm(tot_dl),
                    'Total': minutos_para_hhmm(tot_tt), 'Faturamento': formatar_moeda(tot_fat)
                }
                df_tab = pd.concat([df_tab, pd.DataFrame([total_row])], ignore_index=True)

                def style_row(row):
                    if row['Dia'] == 'TOTAL':
                        return ['background-color:#2a1e0a;color:#f5c842;font-weight:700'] * len(row)
                    return [''] * len(row)

                st.dataframe(
                    df_tab.style.apply(style_row, axis=1),
                    use_container_width=True, hide_index=True, height=340
                )

                legenda = ''.join(
                    f'<span class="legenda-dot"><span style="background:{r["cor"]}"></span>{r["Dia"]}</span>'
                    for _, r in df_resultado.iterrows() if r['Clientes'] > 0
                )
                st.markdown(
                    f'<div class="legenda-row">{legenda}'
                    f'<span class="legenda-dot">🟢 Início &nbsp;🔴 Término</span></div>',
                    unsafe_allow_html=True
                )

            st.session_state.dados_exportar.append({
                'numero': int(rota_num), 'vendedor': nome_vendedor,
                'total_clientes': tot_cli, 'total_km': tot_km,
                'total_atend': tot_at, 'total_desloc': tot_dl, 'total_tempo': tot_tt,
                'fat_rota': fat_rota, 'tabela': df_tab, 'resultado': df_resultado, 'mapa': mapa
            })
            st.markdown("---")

        prog.empty(); stat.empty()
        st.success("Análise concluída.")

    # Exportar
    if st.session_state.dados_exportar:
        st.markdown("### Exportar Relatório")
        c1, c2, c3 = st.columns(3)
        ts = datetime.now().strftime('%Y%m%d_%H%M')
        with c1:
            st.download_button("🗺️ HTML com Mapas",
                gerar_html_relatorio(st.session_state.dados_exportar, True),
                f"poleway_mapas_{ts}.html", "text/html",
                use_container_width=True, type="primary")
        with c2:
            st.download_button("📄 HTML sem Mapas",
                gerar_html_relatorio(st.session_state.dados_exportar, False),
                f"poleway_{ts}.html", "text/html", use_container_width=True)
        with c3:
            st.info("Ctrl+P no navegador para exportar PDF")

    elif not analisar:
        st.markdown("### Dados Carregados")
        ca, cb = st.columns(2)
        with ca:
            st.markdown(f"**Clientes:** `{ARQUIVO_CLIENTES}`")
            st.markdown(f"**Total:** {len(df_clientes)} registros")
            st.markdown(f"**Rotas:** {rotas_disponiveis}")
        with cb:
            st.markdown(f"**Vendedores:** {len(df_vendedores)}")
            st.markdown(f"**Faturamento:** {len(df_fat_rotas)} rotas" if not df_fat_rotas.empty else "**Faturamento:** não carregado")
            st.markdown(f"**Pernoites:** {len(pernoites_data)} rotas" if pernoites_data else "**Pernoites:** não carregado")

if __name__ == "__main__":
    main()