from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pandas as pd
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import datetime
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from openpyxl import Workbook, load_workbook
import win32com.client as win32
from pathlib import Path
import os
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import funcoes as f # Assumido que funcoes.py não contém credenciais hardcoded
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
import webbrowser
import config # Importa o arquivo de configurações sensíveis (NÃO ENVIAR AO GIT!)


# Caminho universal (Anonimizado e seguro)
downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")

# Use SEMPRE essas variáveis nos caminhos dos arquivos:
ordens_path = os.path.join(downloads_path, "ordens_servico_liberadas.xls") 
empresas_path = os.path.join(downloads_path, "dados_empresas.xls")
saida_path = os.path.join(downloads_path, "relatorio_enderecos_saida.xlsx")
saida_csv = os.path.join(downloads_path, "dados_geolocalizacao_saida.csv")
csv_path = saida_csv

# --- FUNÇÕES DE LIMPEZA E UTILS ---

def limpar_arquivos_antigos():
    """Remove arquivos XLS temporários para evitar conflitos de download."""
    for path in [ordens_path, empresas_path]:
        if os.path.exists(path):
            try:
                os.remove(path)
            except PermissionError:
                print(f"ATENÇÃO: Não foi possível remover {os.path.basename(path)}. Feche o arquivo e tente novamente.")

def iniciar_navegador_e_logar(url_sensivel):
    """Inicia o navegador, faz o login e retorna o objeto navegador."""
    # A função f.conexao() deve iniciar o driver sem credenciais.
    navegador = f.conexão() 
    
    # Navega para o site e pede o login manual (anonimizado)
    navegador.get(url_sensivel)
    input("Pressione ENTER no console após o LOGIN manual...") 
    return navegador, ActionChains(navegador)

# --- EXECUÇÃO PRINCIPAL ---

limpar_arquivos_antigos()

# Configuração de datas
data_final = datetime.date.today().strftime('%d/%m/%Y')
data_inicial = str('01/01/2000')

# 1. AUTOMAÇÃO - DOWNLOAD DE DADOS DE EMPRESAS/CLIENTES

# Site da empresa anonimizado (puxado de config.py)
navegador_emp, action_emp = iniciar_navegador_e_logar(config.ARKMEDS_URL_EMPRESA) 

# Clique no botão Data Analyze (ou o botão que inicia a exportação)
WebDriverWait(navegador_emp, 20).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/div/div[1]/div/a[2]'))
).click()

# Abre o dropdown de Empresas
WebDriverWait(navegador_emp, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="div_id_empresas"]/div/span/div/button/b'))
).click()

# Clica na opção 'Todos'
opcao_todos = WebDriverWait(navegador_emp, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Todos')]"))
)
opcao_todos.click()
action_emp.send_keys(Keys.TAB).perform()

# Abre o dropdown das colunas opcionais
WebDriverWait(navegador_emp, 10).until(
    EC.element_to_be_clickable((By.XPATH,'//*[@id="div_id_colunas_opcionais"]/div/span/div/button'))
).click()

# Clica nas opções de colunas (Nome, Endereço)
for texto in ['Nome', 'Endereço']:
    opcao_coluna = WebDriverWait(navegador_emp, 10).until(
        EC.element_to_be_clickable((
            By.XPATH,
            f"//*[@id='div_id_colunas_opcionais']//label[contains(., '{texto}')]"
        ))
    )
    opcao_coluna.click()
    time.sleep(0.2) 

action_emp.send_keys(Keys.TAB).perform()

# Exportar para Excel
WebDriverWait(navegador_emp, 10).until(
    EC.element_to_be_clickable((By.XPATH,'//*[@id="main-content-wrapper"]/div/div[2]/form/div[2]/button'))
).click()

# Aguarda o download de forma segura.
time.sleep(5) 
navegador_emp.quit() 

# 2. AUTOMAÇÃO - DOWNLOAD DE DADOS DE ORDENS DE SERVIÇO (OS)

# Nova instância do navegador para o segundo download
navegador_os, action_os = iniciar_navegador_e_logar(config.ARKMEDS_URL_OS) 

# Clique no botão Data Analyze
WebDriverWait(navegador_os, 20).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="data_analyze_icon"]'))
).click()

# Espera abrir a nova janela e troca para ela
WebDriverWait(navegador_os, 10).until(EC.number_of_windows_to_be(2))
navegador_os.switch_to.window(navegador_os.window_handles[1])

# Abre o dropdown do Estado
WebDriverWait(navegador_os, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="div_id_estado"]/div/span/div/button/b'))
).click()

# Clica na opção 'Liberado Para Entrega'
opcao = WebDriverWait(navegador_os, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Liberado Para Entrega')]"))
)
opcao.click()
action_os.send_keys(Keys.TAB).perform()

# Intervalo de datas
input_intervalo = WebDriverWait(navegador_os, 10).until(
    EC.element_to_be_clickable((By.XPATH,'//*[@id="id_intervalo"]'))
)
input_intervalo.clear()
input_intervalo.send_keys(data_inicial + ' - ' + data_final)
action_os.send_keys(Keys.TAB).perform()

# Abre o dropdown das colunas opcionais
WebDriverWait(navegador_os, 10).until(
    EC.element_to_be_clickable((By.XPATH,'//*[@id="div_id_colunas_opcionais"]/div/span/div/button'))
).click()

# Clica nas opções desejadas
for texto in ['SOLICITANTE', 'TIPO DE EQUIPAMENTO', 'NÚMERO DE SÉRIE', 'MODELO']:
    opcao_coluna = WebDriverWait(navegador_os, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//label[contains(., '{texto}')]"))
    )
    opcao_coluna.click()
    time.sleep(0.2) 

action_os.send_keys(Keys.TAB).perform()

# Exportar para Excel
WebDriverWait(navegador_os, 10).until(
    EC.element_to_be_clickable((By.XPATH,'//*[@id="main-content-wrapper"]/div/div[2]/form/div[2]/button'))
).click()

# Aguarda o download, de forma mais longa por se tratar de um relatório maior
time.sleep(15)
navegador_os.quit()

# 3. PROCESSAMENTO DE DADOS (PANDAS)

ordens_df = pd.read_excel(ordens_path)
empresas_df = pd.read_excel(empresas_path)

resultado = pd.merge(
    ordens_df,
    empresas_df,
    left_on='SOLICITANTE',
    right_on='NOME',
    how='left' 
)

# Seleciona as colunas necessárias para o mapeamento
colunas_selecionadas = [
    'SOLICITANTE', 'CIDADE', 'RUA', 'NUMERO', 'BAIRRO', 
    'TIPO DE EQUIPAMENTO', 'NÚMERO DE SÉRIE', 'MODELO'
]
resultado_final = resultado.reindex(columns=colunas_selecionadas, fill_value='')

# Limpeza e criação da coluna de endereço completo
for col in ['RUA', 'NUMERO', 'BAIRRO', 'CIDADE']:
    resultado_final[col] = resultado_final[col].fillna('')

resultado_final['Endereço Completo'] = (
    resultado_final['RUA'].astype(str) + ', ' +
    resultado_final['NUMERO'].astype(str) + ', ' +
    resultado_final['BAIRRO'].astype(str) + ', ' +
    resultado_final['CIDADE'].astype(str)
)

resultado_final['Endereço Completo'] = resultado_final['Endereço Completo'].str.replace(r'(, )+', ', ', regex=True)
resultado_final['Endereço Completo'] = resultado_final['Endereço Completo'].str.strip(', ').str.strip()

resultado_final.to_excel(saida_path, index=False)
resultado_final.to_csv(saida_csv, index=False, encoding='utf-8-sig')

print(f"\nArquivo CSV pronto para importação/geolocalização: {os.path.basename(saida_csv)}")


# 4. GEOLOCALIZAÇÃO E MAPEAMENTO (FOLIUM)

# Ajuste do certificado (boas práticas)
import certifi
os.environ['SSL_CERT_FILE'] = certifi.where()

df = pd.read_csv(csv_path)

for col in ['SOLICITANTE', 'RUA', 'NUMERO', 'BAIRRO', 'CIDADE']:
    df[col] = df[col].fillna('')

df['EnderecoCompleto'] = (
    df['SOLICITANTE'].astype(str) + ', ' +
    df['RUA'].astype(str) + ', ' +
    df['NUMERO'].astype(str) + ', ' +
    df['BAIRRO'].astype(str) + ', ' +
    df['CIDADE'].astype(str)
)
df['EnderecoCompleto'] = df['EnderecoCompleto'].str.replace(r'(, )+', ', ', regex=True)
df['EnderecoCompleto'] = df['EnderecoCompleto'].str.strip(', ').str.strip()


geolocator = Nominatim(user_agent="mapa_rotas_portfolio", timeout=10) # User agent anonimizado
coordenadas = []

for idx, row in df.iterrows():
    endereco = row['EnderecoCompleto']
    cidade = row['CIDADE']
    resultado = None
    try:
        if endereco and endereco != ', , , ,':
            resultado = geolocator.geocode(f"{endereco}, Brasil")
        if not resultado and cidade:
            resultado = geolocator.geocode(f"{cidade}, Brasil")
        if resultado:
            coordenadas.append((resultado.latitude, resultado.longitude))
        else:
            coordenadas.append((None, None))
    except Exception as e:
        coordenadas.append((None, None))
    time.sleep(1) # Reduzido para 1 segundo para evitar bloqueio da API

df['Latitude'] = [lat for lat, lon in coordenadas]
df['Longitude'] = [lon for lat, lon in coordenadas]
df = df.dropna(subset=['Latitude', 'Longitude'])

if df.empty:
    print("Nenhum endereço foi geocodificado com sucesso. O mapa não será gerado.")
    exit()

m = folium.Map(location=[df.iloc[0]['Latitude'], df.iloc[0]['Longitude']], zoom_start=10)
marker_cluster = MarkerCluster().add_to(m)

for idx, row in df.iterrows():
    popup_text = (
        f"<b>{row.get('SOLICITANTE', '')}</b><br>"
        f"<b>Endereço:</b> {row.get('CIDADE', '')}<br>"
        f"<b>RUA:</b> {row.get('RUA', '')}<br>"
        f"<b>NÚMERO:</b> {row.get('NUMERO', '')}<br>"
        f"<b>BAIRRO:</b> {row.get('BAIRRO', '')}<br>"
        f"<b>Tipo de Equipamento:</b> {row.get('TIPO DE EQUIPAMENTO', '')}<br>"
        f"<b>Modelo:</b> {row.get('MODELO', '')}<br>"
    )
    folium.Marker(
        [row['Latitude'], row['Longitude']],
        popup=folium.Popup(popup_text, max_width=350),
        icon=folium.Icon(color='blue', icon='info-sign')
    ).add_to(marker_cluster)

mapa_html = f"mapa_interativo_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
m.save(mapa_html)
webbrowser.open(mapa_html) 
print(f"Mapa gerado e aberto: {mapa_html}")