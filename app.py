"""
PoleWay - ANALISADOR DE ROTAS DINÂMICO (OSRM Version)
"""

import pandas as pd
import numpy as np
import streamlit as st
import folium
import requests
from scipy.spatial.distance import euclidean # type: ignore
from streamlit_folium import st_folium # type: ignore
import os
import time
from datetime import datetime
import base64
from io import BytesIO
import polyline # type: ignore

# Configs

st.set_page_config(
    page_title="PoleWay - Análise por Rota (OSRM)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dados de Entrada
ARQUIVO_CLIENTES = "rotas_processadas_319.xlsx"
ARQUIVO_TEMPO_ATENDIMENTO = "segmento.xlsx"
ARQUIVO_VENDEDORES = "Vendedores Lat e Long.xlsx"
ARQUIVO_FATURAMENTO = "faturamento.xlsx"
COLUNA_ROTA = "Rota"

# OSRM API
OSRM_BASE_URL = "https://router.project-osrm.org"

# CSS
st.markdown("""
<style>
    .stApp { background: #ffffff !important; font-family: 'Segoe UI', sans-serif; }
    [data-testid="stSidebar"] { background-color: #2b2b2b !important; }
    [data-testid="stSidebar"] * { color: #ffffff !important; }
    h1, h2, h3, p { color: #1d1d1f !important; }
    .stButton > button { background: #FF6B00 !important; color: white !important; }
    
    /* Metrics em preto */
    [data-testid="stMetricValue"] { color: #000000 !important; }
    [data-testid="stMetricLabel"] { color: #000000 !important; }
    
    .rota-header {
        background: linear-gradient(135deg, #D50037, #FF6B00);
        color: white !important;
        padding: 12px 20px;
        border-radius: 8px;
        font-weight: 700;
        font-size: 1.2rem;
        margin: 20px 0 10px 0;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

CITY_COLORS = [
    '#FF0000',
    '#0000FF',
    '#00CC00',
    '#FF00FF',
    '#FF8C00',
    '#00CCCC',
    '#8B4513',
    '#800080',
    '#FFD700',
    '#000000',
    '#DC143C',
    '#1E90FF',
    '#32CD32',
    '#FF1493',
    '#FF4500',
]

# Funções API

def testar_osrm_api():
    try:
        url = f"{OSRM_BASE_URL}/route/v1/driving/-39.3153,-7.2131;-39.3253,-7.2231"
        params = {'overview': 'false'}
        
        response = requests.get(url, params=params, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            return data.get('code') == 'Ok'
        return False
    except:
        return False


def calcular_distancia_osrm(origem, destino):
    try:
        origem_str = f"{origem[1]},{origem[0]}"
        destino_str = f"{destino[1]},{destino[0]}"
        
        url = f"{OSRM_BASE_URL}/route/v1/driving/{origem_str};{destino_str}"
        params = {
            'overview': 'false',
            'steps': 'false'
        }
        
        response = requests.get(url, params=params, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            
            if data.get('code') == 'Ok' and 'routes' in data and len(data['routes']) > 0:
                route = data['routes'][0]
                distancia_m = route.get('distance', 0)  # metros
                tempo_s = route.get('duration', 0)  # segundos
                
                return distancia_m / 1000, tempo_s / 60  # km, minutos        
        return calcular_distancia_euclidiana(origem, destino)
        
    except Exception as e:
        return calcular_distancia_euclidiana(origem, destino)


def calcular_distancia_euclidiana(origem, destino):
    try:
        dist_degrees = euclidean(origem, destino)
        dist_km = dist_degrees * 111 * 1.3
        time_min = (dist_km / 50) * 60
        return dist_km, time_min
    except:
        return 0, 0


def obter_rota_osrm(origem, destino):
    try:
        origem_str = f"{origem[1]},{origem[0]}"
        destino_str = f"{destino[1]},{destino[0]}"
        
        url = f"{OSRM_BASE_URL}/route/v1/driving/{origem_str};{destino_str}"
        params = {
            'overview': 'full',
            'geometries': 'polyline' # Polyline
        }
        
        response = requests.get(url, params=params, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            
            if data.get('code') == 'Ok' and 'routes' in data and len(data['routes']) > 0:
                route = data['routes'][0]
                
                # Decodificar Polyline
                if 'geometry' in route:
                    encoded_polyline = route['geometry']
                    decoded_coords = polyline.decode(encoded_polyline)
                    return [[coord[0], coord[1]] for coord in decoded_coords]
        
        return None
        
    except Exception as e:
        return None


# Funções Auxiliares

def normalizar_cidade(nome):
    if pd.isna(nome):
        return 'DESCONHECIDA'
    nome = str(nome).upper().strip()
    replacements = {
        'Á':'A', 'À':'A', 'Ã':'A', 'Â':'A', 
        'É':'E', 'Ê':'E', 'Í':'I', 
        'Ó':'O', 'Õ':'O', 'Ô':'O', 'Ú':'U', 'Ç':'C'
    }
    for old, new in replacements.items():
        nome = nome.replace(old, new)
    return nome


def detectar_coluna(df, possiveis_nomes):
    for col in possiveis_nomes:
        if col in df.columns:
            return col
    for col in df.columns:
        for nome in possiveis_nomes:
            if nome.lower() in col.lower():
                return col
    return None


def formatar_moeda(valor):
    if pd.isna(valor) or valor == 0:
        return "R$ 0"
    return f"R$ {valor:,.0f}".replace(',', 'X').replace('.', ',').replace('X', '.')


@st.cache_data
def load_faturamento():
    try:
        if not os.path.exists(ARQUIVO_FATURAMENTO):
            return pd.DataFrame(), pd.DataFrame()
        df_rotas = pd.read_excel(ARQUIVO_FATURAMENTO, sheet_name='Rotas', engine='openpyxl')        
        df_cidades = pd.read_excel(ARQUIVO_FATURAMENTO, sheet_name='Cidades', engine='openpyxl')
        
        if 'Rota' in df_rotas.columns:
            df_rotas['Rota'] = pd.to_numeric(df_rotas['Rota'], errors='coerce')
        
        if 'Faturamento' in df_rotas.columns:
            df_rotas['Faturamento'] = pd.to_numeric(df_rotas['Faturamento'], errors='coerce')
        
        if 'Rota' in df_cidades.columns:
            df_cidades['Rota'] = pd.to_numeric(df_cidades['Rota'], errors='coerce')
        
        if 'Faturamento' in df_cidades.columns:
            df_cidades['Faturamento'] = pd.to_numeric(df_cidades['Faturamento'], errors='coerce')
        
        # Normalizar nomes de cidades
        if 'Cidade' in df_cidades.columns:
            df_cidades['Cidade'] = df_cidades['Cidade'].apply(normalizar_cidade)
        
        return df_rotas, df_cidades
        
    except Exception as e:
        st.warning(f"⚠️ Erro ao carregar faturamento: {e}")
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data
def load_vendedores():
    try:
        if not os.path.exists(ARQUIVO_VENDEDORES):
            return pd.DataFrame()
        
        df = pd.read_excel(ARQUIVO_VENDEDORES, engine='openpyxl')
        
        col_lat = detectar_coluna(df, ['latitude', 'Latitude', 'lat', 'Lat'])
        col_lon = detectar_coluna(df, ['longitude', 'Longitude', 'lon', 'Lon', 'lng', 'Lng'])
        col_rota = detectar_coluna(df, ['rota', 'Rota', 'ROTA'])
        
        if not col_lat or not col_lon:
            return pd.DataFrame()
        
        df = df.rename(columns={col_lat: 'latitude', col_lon: 'longitude'})
        if col_rota:
            df = df.rename(columns={col_rota: 'Rota'})
        
        df = df[df['latitude'].notna() & df['longitude'].notna()].copy()
        
        return df
        
    except Exception as e:
        return pd.DataFrame()


@st.cache_data
def load_clientes():
    try:
        if not os.path.exists(ARQUIVO_CLIENTES):
            st.error(f"❌ Arquivo não encontrado: {ARQUIVO_CLIENTES}")
            return pd.DataFrame()
        
        df = pd.read_excel(ARQUIVO_CLIENTES, engine='openpyxl')
        
        if 'latitude' not in df.columns or 'longitude' not in df.columns:
            st.error("❌ Planilha não contém colunas 'latitude' e 'longitude'")
            return pd.DataFrame()
        
        df = df[df['latitude'].notna() & df['longitude'].notna()].copy()
        
        col_rota = detectar_coluna(df, [COLUNA_ROTA, 'rota', 'Rota','ROTA'])
        if col_rota:
            df['ROTA'] = df[col_rota]
        else:
            st.error(f"❌ Coluna '{COLUNA_ROTA}' não encontrada!")
            return pd.DataFrame()
        
        col_cidade = detectar_coluna(df, ['cidade', 'Cidade', 'dsCidadeComercial', 'nmCidade', 'municipio'])
        if col_cidade:
            df['Cidade'] = df[col_cidade].apply(normalizar_cidade)
        else:
            if 'endereco_completo' in df.columns:
                df['Cidade'] = df['endereco_completo'].apply(
                    lambda x: normalizar_cidade(str(x).split(',')[-2].strip()) 
                    if pd.notna(x) and len(str(x).split(',')) >= 2 else 'DESCONHECIDA'
                )
            else:
                df['Cidade'] = 'DESCONHECIDA'
        
        df = carregar_tempo_atendimento(df)
        
        return df
        
    except Exception as e:
        st.error(f"Erro ao carregar clientes: {e}")
        return pd.DataFrame()

def carregar_tempo_atendimento(df):
    if not os.path.exists(ARQUIVO_TEMPO_ATENDIMENTO):
        df['TempoVisita'] = 30
        return df
    
    try:
        df_tempo = pd.read_excel(ARQUIVO_TEMPO_ATENDIMENTO, engine='openpyxl')
        
        col_segmento_tempo = None
        col_tempo_visita = None
        
        for col in df_tempo.columns:
            if 'segmento' in col.lower():
                col_segmento_tempo = col
            if 'tempo' in col.lower() and 'visita' in col.lower():
                col_tempo_visita = col
        
        if col_segmento_tempo and col_tempo_visita:
            df_tempo['Segmento_Key'] = df_tempo[col_segmento_tempo].astype(str).str.strip().str.upper()
            df_tempo['TempoVisita_Val'] = pd.to_numeric(df_tempo[col_tempo_visita], errors='coerce').fillna(30)
            
            col_segmento_cli = None
            for col in df.columns:
                if 'segmento' in col.lower():
                    col_segmento_cli = col
                    break
            
            if col_segmento_cli:
                df['Segmento_Key'] = df[col_segmento_cli].astype(str).str.strip().str.upper()
                tempo_map = dict(zip(df_tempo['Segmento_Key'], df_tempo['TempoVisita_Val']))
                df['TempoVisita'] = df['Segmento_Key'].map(tempo_map).fillna(30)
                df = df.drop(columns=['Segmento_Key'], errors='ignore')
            else:
                df['TempoVisita'] = 30
        else:
            df['TempoVisita'] = 30
            
    except Exception as e:
        df['TempoVisita'] = 30
    
    return df


def order_nearest_neighbor(clientes, start):
    if not clientes:
        return []
    
    unvisited = clientes.copy()
    ordered = []
    current = start
    
    while unvisited:
        nearest = min(unvisited, key=lambda c: euclidean(current, (c['latitude'], c['longitude'])))
        ordered.append(nearest)
        unvisited.remove(nearest)
        current = (nearest['latitude'], nearest['longitude'])
    
    return ordered


def minutos_para_hhmm(minutos):
    if pd.isna(minutos) or minutos == 0:
        return "00:00"
    horas = int(minutos // 60)
    mins = int(minutos % 60)
    return f"{horas:02d}:{mins:02d}"

# Funções de Rota
def analisar_rota(df_rota, rota_num, vendor_home, df_faturamento_cidades, usar_osrm=True, progress_callback=None):    
    if df_rota.empty:
        return None, None
    
    cidades_da_rota = sorted(df_rota['Cidade'].dropna().unique().tolist())
    
    col_cliente = detectar_coluna(df_rota, ['cdCliente', 'codigo', 'Codigo', 'Cliente', 'id'])
    col_nome = detectar_coluna(df_rota, ['nmFantasia', 'nome', 'Nome', 'razao_social'])
    
    resultados = []
    todos_clientes = []
    
    for idx, cidade in enumerate(cidades_da_rota):
        df_cidade = df_rota[df_rota['Cidade'] == cidade].copy()
        
        if df_cidade.empty:
            continue
        
        clientes = df_cidade.to_dict('records')
        clientes_ordenados = order_nearest_neighbor(clientes, vendor_home)
        
        km_total = 0
        tempo_atend = 0
        tempo_desloc = 0
        
        cor = CITY_COLORS[idx % len(CITY_COLORS)]
        
        # Buscar faturamento da cidade
        faturamento_cidade = 0
        if not df_faturamento_cidades.empty:
            fat_filtro = df_faturamento_cidades[
                (df_faturamento_cidades['Rota'] == rota_num) & 
                (df_faturamento_cidades['Cidade'] == cidade)
            ]
            if not fat_filtro.empty:
                faturamento_cidade = fat_filtro['Faturamento'].iloc[0]
        
        for i, cli in enumerate(clientes_ordenados):
            cli['cor'] = cor
            cli['cidade_display'] = cidade.title()
            cli['ordem'] = i + 1
            cli['cod_cliente'] = cli.get(col_cliente, '') if col_cliente else ''
            cli['nome_cliente'] = cli.get(col_nome, '') if col_nome else ''
            
            tempo_visita = cli.get('TempoVisita', 30)
            if pd.isna(tempo_visita) or tempo_visita == 0:
                tempo_visita = 30
            tempo_atend += float(tempo_visita)
            
            if i < len(clientes_ordenados) - 1:
                origem = (cli['latitude'], cli['longitude'])
                destino = (clientes_ordenados[i+1]['latitude'], clientes_ordenados[i+1]['longitude'])
                
                if usar_osrm:
                    dist, tempo = calcular_distancia_osrm(origem, destino)
                else:
                    dist, tempo = calcular_distancia_euclidiana(origem, destino)
                
                km_total += dist
                tempo_desloc += tempo
        
        resultados.append({
            'Cidades': cidade.title(),
            'Clientes': len(clientes_ordenados),
            'KM': round(km_total, 1),
            'Atend.': int(round(tempo_atend)),
            'Desloc.': int(round(tempo_desloc)),
            'Total': int(round(tempo_atend + tempo_desloc)),
            'Faturamento': faturamento_cidade,
            'cor': cor
        })
        
        todos_clientes.extend(clientes_ordenados)
    
    df_resultado = pd.DataFrame(resultados)
    
    return df_resultado, todos_clientes

# Mapa
def criar_mapa(clientes, vendor_home, rota_num, nome_vendedor="", usar_rotas_osrm=True):
    
    if not clientes:
        m = folium.Map(location=[-7.2, -39.3], zoom_start=9, tiles='CartoDB positron')
        return m
    
    lats = [c['latitude'] for c in clientes] + [vendor_home[0]]
    lons = [c['longitude'] for c in clientes] + [vendor_home[1]]
    center = [sum(lats)/len(lats), sum(lons)/len(lons)]
    
    m = folium.Map(location=center, zoom_start=10, tiles='CartoDB positron')
    
    folium.Marker(
        location=vendor_home,
        popup=f"🏠 <b>Vendedor Rota {rota_num}</b><br>{nome_vendedor}",
        tooltip=f"Vendedor - {nome_vendedor}",
        icon=folium.Icon(color='darkblue', icon='home', prefix='fa')
    ).add_to(m)
    
    por_cidade = {}
    for c in clientes:
        cidade = c.get('cidade_display', 'N/A')
        if cidade not in por_cidade:
            por_cidade[cidade] = []
        por_cidade[cidade].append(c)
    
    for cidade, lista in por_cidade.items():
        lista = sorted(lista, key=lambda x: x.get('ordem', 0))
        cor = lista[0].get('cor', '#999')
        
        if len(lista) > 1:
            for i in range(len(lista) - 1):
                origem = lista[i]
                destino = lista[i + 1]
                
                origem_coords = (origem['latitude'], origem['longitude'])
                destino_coords = (destino['latitude'], destino['longitude'])
                
                rota_coords = None
                
                if usar_rotas_osrm:
                    rota_coords = obter_rota_osrm(origem_coords, destino_coords)
                
                if rota_coords and len(rota_coords) > 1:
                    folium.PolyLine(
                        rota_coords, 
                        color=cor, 
                        weight=5, 
                        opacity=0.9,
                        tooltip=f"{cidade}: {origem.get('cod_cliente', '')} → {destino.get('cod_cliente', '')}"
                    ).add_to(m)
                else:
                    folium.PolyLine(
                        [[origem['latitude'], origem['longitude']], 
                         [destino['latitude'], destino['longitude']]], 
                        color=cor, 
                        weight=3, 
                        opacity=0.5,
                        dash_array='5, 10',
                        tooltip=f"{cidade} (linha reta)"
                    ).add_to(m)
        
        for cli in lista:
            codigo = cli.get('cod_cliente', '')
            nome = cli.get('nome_cliente', '')
            ordem = cli.get('ordem', '')
            
            popup_html = f"""
            <div style="min-width: 200px;">
                <h4 style="color: {cor}; margin: 0 0 8px 0;">🏷️ {codigo}</h4>
                <p style="margin: 2px 0;"><b>Nome:</b> {nome}</p>
                <p style="margin: 2px 0;"><b>Cidade:</b> {cidade}</p>
                <p style="margin: 2px 0;"><b>Ordem visita:</b> {ordem}</p>
                <hr>
                <p style="font-size: 10px; color: #666;">
                    Lat: {cli['latitude']:.6f}<br>
                    Lon: {cli['longitude']:.6f}
                </p>
            </div>
            """
            
            folium.CircleMarker(
                location=[cli['latitude'], cli['longitude']],
                radius=8,
                color=cor,
                fill=True,
                fill_color=cor,
                fill_opacity=0.9,
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=f"{ordem}. {codigo} - {cidade}"
            ).add_to(m)
    
    return m


# Exportar Reltório HTML
def gerar_html_relatorio(dados_rotas, incluir_mapas=True):  
    data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    # Calcular totais gerais
    qtd_rotas = len(dados_rotas)
    total_geral_clientes = sum(r['total_clientes'] for r in dados_rotas) / qtd_rotas if qtd_rotas > 0 else 0
    total_geral_km = sum(r['total_km'] for r in dados_rotas) / qtd_rotas if qtd_rotas > 0 else 0
    total_geral_tempo = sum(r['total_tempo'] for r in dados_rotas) / qtd_rotas if qtd_rotas > 0 else 0
    
    # Calcular média geral de faturamento
    soma_faturamentos = sum(r.get('media_faturamento_rota', 0) for r in dados_rotas)
    media_geral_faturamento = soma_faturamentos / qtd_rotas if qtd_rotas > 0 else 0
    
    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PoleWay - Relatório de Rotas (OSRM)</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            color: #333;
            line-height: 1.6;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }}
        
        .header {{
            background: linear-gradient(135deg, #D50037, #FF6B00);
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2rem;
            margin-bottom: 10px;
        }}
        
        .header .subtitle {{
            font-size: 1rem;
            opacity: 0.9;
        }}
        
        .resumo-geral {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .card-resumo {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            text-align: center;
        }}
        
        .card-resumo .valor {{
            font-size: 2rem;
            font-weight: bold;
            color: #D50037;
        }}
        
        .card-resumo .label {{
            font-size: 0.9rem;
            color: #666;
            margin-top: 5px;
        }}
        
        .rota-section {{
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 30px;
            overflow: hidden;
        }}
        
        .rota-header {{
            background: linear-gradient(135deg, #D50037, #FF6B00);
            color: white;
            padding: 15px 20px;
            font-size: 1.2rem;
            font-weight: bold;
        }}
        
        .rota-content {{
            padding: 20px;
        }}
        
        .rota-info {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 8px;
        }}
        
        .rota-info-item {{
            text-align: center;
        }}
        
        .rota-info-item .valor {{
            font-size: 1.5rem;
            font-weight: bold;
            color: #333;
        }}
        
        .rota-info-item .label {{
            font-size: 0.8rem;
            color: #666;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }}
        
        th, td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #eee;
        }}
        
        th {{
            background: #f8f9fa;
            font-weight: 600;
            color: #333;
        }}
        
        tr:hover {{
            background: #f8f9fa;
        }}
        
        tr.total-row {{
            background: #FFC629 !important;
            font-weight: bold;
        }}
        
        tr.total-row td {{
            border-bottom: none;
        }}
        
        .legenda {{
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            margin-top: 15px;
            padding-top: 15px;
            border-top: 1px solid #eee;
        }}
        
        .legenda-item {{
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 0.9rem;
        }}
        
        .legenda-cor {{
            width: 14px;
            height: 14px;
            border-radius: 50%;
        }}
        
        .rota-layout {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 20px;
        }}
        
        .mapa-container {{
            height: 400px;
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid #ddd;
        }}
        
        .mapa-container iframe {{
            width: 100%;
            height: 100%;
            border: none;
        }}
        
        .tabela-container {{
            overflow-x: auto;
        }}
        
        @media (max-width: 900px) {{
            .rota-layout {{
                grid-template-columns: 1fr;
            }}
        }}
        
        .footer {{
            text-align: center;
            padding: 20px;
            color: #666;
            font-size: 0.85rem;
        }}
        
        .osrm-badge {{
            background: #4CAF50;
            color: white;
            padding: 5px 10px;
            border-radius: 5px;
            font-size: 0.85rem;
            display: inline-block;
            margin-left: 10px;
        }}
        
        @media print {{
            body {{
                background: white;
            }}
            
            .rota-section {{
                break-inside: avoid;
                page-break-inside: avoid;
            }}
            
            .header {{
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            
            .rota-header {{
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            
            tr.total-row {{
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 PoleWay - Relatório de Rotas</h1>
            <div class="subtitle">
                Gerado em {data_geracao}
                <span class="osrm-badge">🌍 Powered by OSRM</span>
            </div>
        </div>
        
        <div class="resumo-geral">
            <div class="card-resumo">
                <div class="valor">{qtd_rotas}</div>
                <div class="label">Rotas Analisadas</div>
            </div>
            <div class="card-resumo">
                <div class="valor">{total_geral_clientes:.0f}</div>
                <div class="label">Média de Clientes</div>
            </div>
            <div class="card-resumo">
                <div class="valor">{total_geral_km:.1f} km</div>
                <div class="label">Distância Média</div>
            </div>
            <div class="card-resumo">
                <div class="valor">{minutos_para_hhmm(total_geral_tempo)}</div>
                <div class="label">Tempo Médio</div>
            </div>
            <div class="card-resumo">
                <div class="valor">{formatar_moeda(media_geral_faturamento)}</div>
                <div class="label">Faturamento Médio</div>
            </div>
        </div>
"""
    
    # Adicionar cada rota
    for rota in dados_rotas:
        # Gerar HTML do mapa se disponível
        mapa_html = ""
        if incluir_mapas and 'mapa' in rota and rota['mapa'] is not None:
            try:
                mapa_html = rota['mapa']._repr_html_()
            except:
                mapa_html = ""
        
        html += f"""
        <div class="rota-section">
            <div class="rota-header">
                🚚 Rota {rota['numero']} - {rota['total_clientes']} clientes | 👤 {rota['vendedor']} | 💰 {formatar_moeda(rota.get('media_faturamento_rota', 0))}
            </div>
            <div class="rota-content">
                <div class="rota-info">
                    <div class="rota-info-item">
                        <div class="valor">{rota['total_clientes']}</div>
                        <div class="label">Clientes</div>
                    </div>
                    <div class="rota-info-item">
                        <div class="valor">{rota['total_km']:.1f} km</div>
                        <div class="label">Distância</div>
                    </div>
                    <div class="rota-info-item">
                        <div class="valor">{minutos_para_hhmm(rota['total_atend'])}</div>
                        <div class="label">Atendimento</div>
                    </div>
                    <div class="rota-info-item">
                        <div class="valor">{minutos_para_hhmm(rota['total_desloc'])}</div>
                        <div class="label">Deslocamento</div>
                    </div>
                    <div class="rota-info-item">
                        <div class="valor">{minutos_para_hhmm(rota['total_tempo'])}</div>
                        <div class="label">Tempo Total</div>
                    </div>
                    <div class="rota-info-item">
                        <div class="valor">{formatar_moeda(rota.get('media_faturamento_rota', 0))}</div>
                        <div class="label">Faturamento</div>
                    </div>
                </div>
                
                <div class="rota-layout">
"""
        
        # Adicionar mapa se disponível
        if mapa_html:
            html += f"""
                    <div class="mapa-container">
                        {mapa_html}
                    </div>
"""
        
        # Tabela
        html += """
                    <div class="tabela-container">
                        <table>
                            <thead>
                                <tr>
                                    <th>Cidade</th>
                                    <th>Clientes</th>
                                    <th>KM</th>
                                    <th>Atend.</th>
                                    <th>Desloc.</th>
                                    <th>Total</th>
                                    <th>Faturamento</th>
                                </tr>
                            </thead>
                            <tbody>
"""
        
        # Linhas da tabela
        for _, row in rota['tabela'].iterrows():
            is_total = row['Cidades'] == 'TOTAL'
            row_class = 'total-row' if is_total else ''
            
            html += f"""
                        <tr class="{row_class}">
                            <td>{row['Cidades']}</td>
                            <td>{row['Clientes']}</td>
                            <td>{row['KM']}</td>
                            <td>{row['Atend.']}</td>
                            <td>{row['Desloc.']}</td>
                            <td>{row['Total']}</td>
                            <td>{row['Faturamento']}</td>
                        </tr>
"""
        
        html += """
                            </tbody>
                        </table>
                    </div>
                </div>
                
                <div class="legenda">
"""
        
        # Legenda de cores
        for _, row in rota['resultado'].iterrows():
            if row['Clientes'] > 0:
                html += f"""
                    <div class="legenda-item">
                        <div class="legenda-cor" style="background: {row['cor']};"></div>
                        <span>{row['Cidades']}</span>
                    </div>
"""
        
        html += """
                </div>
            </div>
        </div>
"""
    
    html += """
        <div class="footer">
            <p>PoleWay - Sistema de Análise de Rotas</p>
            <p>Rotas calculadas com OSRM (Open Source Routing Machine)</p>
            <p>Relatório gerado automaticamente</p>
        </div>
    </div>
</body>
</html>
"""
    
    return html


# Front End Interface
def main():
    st.markdown("# 📊 Análise de Rotas Dinâmica (OSRM)")
    st.markdown("**Com OSRM API (Open Source Routing Machine) - Gratuito e Open Source**")
    st.markdown("---")
    
    df_clientes = load_clientes()
    df_vendedores = load_vendedores()
    df_faturamento_rotas, df_faturamento_cidades = load_faturamento()
    
    osrm_ok = testar_osrm_api()
    
    if df_clientes.empty:
        st.error("❌ Erro ao carregar clientes")
        st.stop()
    
    rotas_disponiveis = sorted(df_clientes['ROTA'].dropna().unique().tolist())
    
    with st.sidebar:
        st.markdown("## ⚙️ Configurações")
        
        if osrm_ok:
            st.success("✅ OSRM API OK")
        else:
            st.error("❌ OSRM API não disponível")
        
        if not df_faturamento_rotas.empty and not df_faturamento_cidades.empty:
            st.success(f"✅ Faturamento carregado")
            st.info(f"📊 {len(df_faturamento_rotas)} rotas com dados")
        else:
            st.warning("⚠️ Faturamento não disponível")
        
        st.info("🌍 **OSRM** - Roteamento gratuito e open source")
        
        usar_osrm = st.checkbox("🗺️ Usar OSRM API", value=osrm_ok, 
                                 help="Desmarque para usar cálculo aproximado (mais rápido)")
        
        if usar_osrm and not osrm_ok:
            st.warning("⚠️ OSRM não disponível - usando cálculo aproximado")
        
        st.markdown("### 🗺️ Rotas")
        rotas_selecionadas = st.multiselect(
            "Selecione as rotas:",
            options=rotas_disponiveis,
            default=rotas_disponiveis[:3] if len(rotas_disponiveis) > 3 else rotas_disponiveis,
            format_func=lambda x: f"Rota {int(x)}" if pd.notna(x) else "Sem rota"
        )
        
        st.markdown("---")
        analisar = st.button("📊 Analisar", use_container_width=True, type="primary")
    
    # Inicializar session_state para dados das rotas
    if 'dados_rotas_exportar' not in st.session_state:
        st.session_state.dados_rotas_exportar = []
    
    if analisar:
        st.session_state.dados_rotas_exportar = []
        
        total_clientes = len(df_clientes)
        total_rotas = len(rotas_disponiveis)
        
        if not df_faturamento_rotas.empty:
            rotas_com_fat = df_faturamento_rotas[df_faturamento_rotas['Rota'].isin(rotas_selecionadas)]
            if not rotas_com_fat.empty:
                media_geral_fat = rotas_com_fat['Faturamento'].mean()
            else:
                media_geral_fat = 0
        else:
            media_geral_fat = 0
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📊 Total Clientes", total_clientes)
        col2.metric("🗺️ Total Rotas", total_rotas)
        col3.metric("👤 Vendedores", len(df_vendedores))
        col4.metric("💰 Faturamento Médio", formatar_moeda(media_geral_fat))
        
        st.markdown("---")
        
        # Progress bar
        progress = st.progress(0)
        status = st.empty()
        
        for i, rota_num in enumerate(sorted(rotas_selecionadas)):
            progress.progress((i + 1) / len(rotas_selecionadas))
            status.text(f"Processando Rota {int(rota_num)}...")
            
            df_rota = df_clientes[df_clientes['ROTA'] == rota_num].copy()
            
            if df_rota.empty:
                continue
            
            # Alterar para Coordenada do Vendedor!
            vendor_home = (-7.2131, -39.3153)  
            nome_vendedor = "Não definido"
            
            if not df_vendedores.empty and 'Rota' in df_vendedores.columns:
                vendedor = df_vendedores[df_vendedores['Rota'] == rota_num]
                if not vendedor.empty:
                    vendor_home = (vendedor.iloc[0]['latitude'], vendedor.iloc[0]['longitude'])
                    col_nome = detectar_coluna(vendedor, ['nome', 'Nome', 'vendedor', 'Vendedor'])
                    if col_nome:
                        nome_vendedor = vendedor.iloc[0][col_nome]
            
            media_faturamento_rota = 0
            if not df_faturamento_rotas.empty:
                fat_rota = df_faturamento_rotas[df_faturamento_rotas['Rota'] == rota_num]
                if not fat_rota.empty:
                    media_faturamento_rota = fat_rota['Faturamento'].iloc[0]
            
            df_resultado, clientes = analisar_rota(
                df_rota, rota_num, vendor_home, df_faturamento_cidades, usar_osrm
            )
            
            if df_resultado is None or df_resultado.empty:
                continue
            
            total_clientes_rota = df_resultado['Clientes'].sum()
            if total_clientes_rota == 0:
                continue
            
            # Header
            st.markdown(
                f'<div class="rota-header">🚚 Rota {int(rota_num)} - {total_clientes_rota} clientes | 👤 {nome_vendedor} | 💰 {formatar_moeda(media_faturamento_rota)}</div>', 
                unsafe_allow_html=True
            )
            
            # Layout
            col1, col2 = st.columns([1, 1])
            
            with col1:
                mapa = criar_mapa(clientes, vendor_home, rota_num, nome_vendedor, usar_osrm)
                st_folium(mapa, width=500, height=400, key=f"map_{rota_num}", returned_objects=[])
            
            with col2:
                # Tabela
                df_tabela = df_resultado[['Cidades', 'Clientes', 'KM', 'Atend.', 'Desloc.', 'Total', 'Faturamento']].copy()
                
                # Totais
                total_clientes_tab = df_tabela['Clientes'].sum()
                total_km = df_tabela['KM'].sum()
                total_atend = df_tabela['Atend.'].sum()
                total_desloc = df_tabela['Desloc.'].sum()
                total_geral = df_tabela['Total'].sum()
                total_faturamento = df_tabela['Faturamento'].sum()
                
                # Formatar KM com 1 casa decimal
                df_tabela['KM'] = df_tabela['KM'].apply(lambda x: f"{x:.1f}")

                # Converter para HH:MM
                df_tabela['Atend.'] = df_tabela['Atend.'].apply(minutos_para_hhmm)
                df_tabela['Desloc.'] = df_tabela['Desloc.'].apply(minutos_para_hhmm)
                df_tabela['Total'] = df_tabela['Total'].apply(minutos_para_hhmm)
                
                # Formatar faturamento
                df_tabela['Faturamento'] = df_tabela['Faturamento'].apply(formatar_moeda)
                
                total = {
                    'Cidades': 'TOTAL',
                    'Clientes': total_clientes_tab,
                    'KM': f"{total_km:.1f}",
                    'Atend.': minutos_para_hhmm(total_atend),
                    'Desloc.': minutos_para_hhmm(total_desloc),
                    'Total': minutos_para_hhmm(total_geral),
                    'Faturamento': formatar_moeda(total_faturamento)
                }
                df_tabela = pd.concat([df_tabela, pd.DataFrame([total])], ignore_index=True)
                
                def style_table(row):
                    if row['Cidades'] == 'TOTAL':
                        return ['background-color: #FFC629; font-weight: bold; color: black'] * len(row)
                    return [''] * len(row)
                
                styled = df_tabela.style.apply(style_table, axis=1)
                st.dataframe(styled, use_container_width=True, hide_index=True, height=300)
                
                st.markdown("**Legenda:**")
                cores_html = ""
                for _, row in df_resultado.iterrows():
                    if row['Clientes'] > 0:
                        cores_html += f'<span style="display:inline-block;width:12px;height:12px;background:{row["cor"]};border-radius:50%;margin-right:5px;"></span>{row["Cidades"]} '
                st.markdown(cores_html, unsafe_allow_html=True)
            
            # Salvar dados para exportação
            st.session_state.dados_rotas_exportar.append({
                'numero': int(rota_num),
                'vendedor': nome_vendedor,
                'total_clientes': total_clientes_tab,
                'total_km': total_km,
                'total_atend': total_atend,
                'total_desloc': total_desloc,
                'total_tempo': total_geral,
                'media_faturamento_rota': media_faturamento_rota,
                'tabela': df_tabela,
                'resultado': df_resultado,
                'mapa': mapa
            })
            
            st.markdown("---")
        
        progress.empty()
        status.empty()
        st.success("✅ Análise concluída!")
    
    # Botão de exportar
    if st.session_state.dados_rotas_exportar:
        st.markdown("### 📥 Exportar Relatório")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            html_com_mapas = gerar_html_relatorio(st.session_state.dados_rotas_exportar, incluir_mapas=True)
            
            st.download_button(
                label="🗺️ HTML com Mapas",
                data=html_com_mapas,
                file_name=f"relatorio_rotas_osrm_com_mapas_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                mime="text/html",
                use_container_width=True,
                type="primary"
            )
        
        with col2:
            html_sem_mapas = gerar_html_relatorio(st.session_state.dados_rotas_exportar, incluir_mapas=False)
            
            st.download_button(
                label="📄 HTML sem Mapas",
                data=html_sem_mapas,
                file_name=f"relatorio_rotas_osrm_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                mime="text/html",
                use_container_width=True
            )
        
        with col3:
            st.info("💡 Use **Ctrl+P** no navegador para salvar como PDF")
    
    elif not analisar:
        st.info("👆 Clique em **Analisar** para processar as rotas")
        
        # Infos
        st.markdown("### 📋 Dados Carregados")
        col1, col2 = st.columns(2)
        
        with col1:
            st.write(f"**Clientes:** {ARQUIVO_CLIENTES}")
            st.write(f"**Total:** {len(df_clientes)} registros")
            st.write(f"**Rotas encontradas:** {rotas_disponiveis}")
        
        with col2:
            st.write(f"**Vendedores:** {ARQUIVO_VENDEDORES}")
            if not df_vendedores.empty:
                st.write(f"**Total:** {len(df_vendedores)} vendedores")
            else:
                st.write("⚠️ Nenhum vendedor carregado")
            
            if not df_faturamento_rotas.empty:
                st.write(f"**Faturamento:** {len(df_faturamento_rotas)} rotas")
            else:
                st.write("⚠️ Faturamento não carregado")

if __name__ == "__main__":
    main()