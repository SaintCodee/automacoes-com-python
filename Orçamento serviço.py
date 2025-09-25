# Versão aprimorada para mais de 2 orçamentos.
# Este código foi revisado para remover informações sensíveis e melhorar a segurança.

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar, messagebox
import sys

def cria_OS(prazo_entrega, numero_OS):
    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument("--log-level=3")  # Silencia logs desnecessários

    # Gerenciador para o driver do Edge
    service = Service(EdgeChromiumDriverManager().install())
    navegador = webdriver.Edge(service=service, options=edge_options)
    
    navegador.implicitly_wait(30)
    action = ActionChains(navegador)
    site = "https://aqui_vai_o_site_da_empresa.com.br/ordem_servico/"

    try:
        navegador.get(site)
        print("Navegador aberto. Faça login manualmente no sistema.")
        print("Após o login, não feche o navegador. O script continuará automaticamente.")
        
        # O script agora espera que você faça o login manual e navegue até a página inicial
        WebDriverWait(navegador, 120).until(
            EC.presence_of_element_located((By.ID, "datatable_filter"))
        )
        print("Login detectado. Prosseguindo com a coleta de dados.")

        # Coleta de Dados
        cadastro = []
        for i, ordem in enumerate(numero_OS):
            try:
                # URL genérica para a página de detalhes da OS
                navegador.get(f"aqui_vai_a_url_da_os/{ordem}")
                WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="cliente_nome"]'))
                )
                
                # Acessando elementos com nomes genéricos
                nome_cliente = navegador.find_element(By.XPATH, '//*[@id="cliente_nome"]').text
                tipo_equipamento = navegador.find_element(By.XPATH, '//*[@id="equipamento_tipo"]').text
                marca_equipamento = navegador.find_element(By.XPATH, '//*[@id="equipamento_marca"]').text
                modelo_equipamento = navegador.find_element(By.XPATH, '//*[@id="equipamento_modelo"]').text
                ns_equipamento = navegador.find_element(By.XPATH, '//*[@id="equipamento_ns"]').text
                servicos_equipamento = navegador.find_element(By.XPATH, '//*[@id="servicos_solicitados"]').text
                tipo_servico = "Tipo de Serviço Padrão"
                ident = "Identificação do equipamento: "
                
                cadastro.append([
                    nome_cliente,
                    tipo_equipamento,
                    marca_equipamento,
                    modelo_equipamento,
                    ns_equipamento,
                    servicos_equipamento,
                    tipo_servico,
                    ident
                ])
                print(f"Dados da OS {ordem} coletados com sucesso.")
            except (NoSuchElementException, TimeoutException) as e:
                print(f"Erro ao coletar dados da OS {ordem}: {e}. Pulando esta OS.")
                continue

        # Normalização dos dados
        for dados_os in cadastro:
            dados_os[5] = dados_os[5].replace("Troca", "Substituição").replace("troca", "Substituição") \
                                   .replace("Substituir", "Substituição do").replace("Cotar", "Substituição do") \
                                   .replace("cotar", "Substituição do").replace("Reposição", "Substituição") \
                                   .replace("reposição", "Substituição").replace("Reparo", "Manutenção Corretiva") \
                                   .replace("reparo", "Manutenção Corretiva").replace("Limpeza", "Manutenção Preventiva") \
                                   .replace("limpeza", "Manutenção Preventiva")
            dados_os[5] = dados_os[5].replace(", ", "\n").replace(",", "\n").replace("/", "\n").title()

        # Criação das propostas
        print("Iniciando a criação de orçamentos.")
        for j, Orcamentos in enumerate(numero_OS):
            if j >= len(cadastro) or not cadastro[j]:
                continue
            
            try:
                # URL genérica para a página de criação de orçamentos
                navegador.get("aqui_vai_a_url_de_orcamentos/servicos/")
                
                # Acessando elementos com XPaths genéricos
                WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="select2-orcamento-solicitante-servico-container"]'))
                ).click()
                time.sleep(2)
                action.send_keys(cadastro[j][0]).perform()
                time.sleep(7)
                action.send_keys(Keys.ENTER).perform()
                navegador.find_element(By.XPATH, '//*[@id="comeco_orcamento_servicos"]/div/div[2]/div/div[3]/button[2]').click()
                
                # Segunda etapa: tipo de serviço e equipamento
                WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="form-servicos"]/div/div[1]/div/div/div/span[2]/span[1]/span/span[2]/b'))
                ).click()
                time.sleep(2)
                action.send_keys(cadastro[j][1]).perform()
                time.sleep(20)
                action.send_keys(Keys.ENTER).perform()
                navegador.find_element(By.XPATH, '//*[@id="quantidade-servico"]').click()
                action.send_keys("1").perform()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="botao-servico"]/strong').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="comeco_orcamento_servicos"]/div/div[2]/div/div[3]/button[2]/i').click()
                time.sleep(5)
                
                # Terceira etapa: pagamento
                WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="valor_desconto"]'))
                ).click()
                time.sleep(1)
                action.send_keys(Keys.TAB).perform()
                time.sleep(1)
                action.send_keys("Faturado").perform()
                time.sleep(1)
                action.send_keys(Keys.ENTER).perform()
                time.sleep(1)
                action.send_keys(Keys.TAB).perform()
                time.sleep(1)
                action.send_keys("Outros").perform()
                time.sleep(1)
                action.send_keys(Keys.ENTER).perform()
                action.send_keys(Keys.TAB).perform()
                action.send_keys("Boleto 20 dias").perform()
                navegador.find_element(By.XPATH, '//*[@id="comeco_orcamento_servicos"]/div/div[2]/div/div[3]/button[2]/i').click()
                time.sleep(1)
                
                # Última etapa: descritivo do serviço
                OS = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="numero"]'))
                )
                OS.clear()
                OS.click()
                action.send_keys(str(numero_OS[j])).perform()
                navegador.find_element(By.XPATH, '//*[@id="prazo_entrega"]').click()
                action.send_keys(prazo_entrega).perform()
                navegador.find_element(By.XPATH, '//*[@id="validade"]').click()
                action.send_keys("90").perform()
                time.sleep(0.5)
                action.send_keys(Keys.TAB).perform()
                time.sleep(0.5)
                action.send_keys(Keys.TAB).perform()
                time.sleep(0.5)
                action.send_keys(Keys.TAB).perform()
                action.send_keys(f"Equipamento: {cadastro[j][1]} {cadastro[j][2]}, Modelo: {cadastro[j][3]}, Número de Série: {cadastro[j][4]}, {cadastro[j][7]}").perform()
                action.send_keys("\n\nServiços a serem executados:\n").perform()
                action.send_keys(cadastro[j][5]).perform()
                action.send_keys("\nLimpeza externa do equipamento\nTestes de Funcionamento\nAjustes Finais\n\nObs: A garantia não cobre peças não substituídas, mau uso ou serviços que não foram realizados.").perform()
                
                WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="form4"]/div[2]/div[8]/div[1]/div/span/span[1]/span/span[2]'))
                ).click()
                # Nome do colaborador foi substituído por um nome fictício
                action.send_keys("Colaborador Padrão").perform()
                time.sleep(5)
                action.send_keys(Keys.ENTER).perform()
                navegador.find_element(By.XPATH, '//*[@id="form4"]/div[2]/div[8]/div[2]/div/div').click()
                print(f"Orçamento para a OS {Orcamentos} criado com sucesso.")
            
            except (NoSuchElementException, TimeoutException) as e:
                print(f"Erro ao criar orçamento para a OS {Orcamentos}: {e}")
                print("Continuando para a próxima OS...")
                continue
                
    except Exception as e:
        print(f"Ocorreu um erro crítico: {e}")
        messagebox.showerror("Erro Crítico", f"Ocorreu um erro: {e}\nO programa será encerrado.")
    finally:
        print("Finalizando o programa. Fechando o navegador...")
        navegador.quit()
        sys.exit()

# Funções da GUI (interface gráfica) permanecem as mesmas
def adicionar_valor():
    valor = N_Ordem.get()
    if valor:
        numero_OS.append(valor)
        N_Ordem.set('')
        lista_valores.set(','.join(numero_OS))
    else:
        messagebox.showwarning("Aviso", "Por favor, insira um número de OS.")

def prosseguir():
    prazo = prazo_entrega.get()
    if not numero_OS:
        messagebox.showwarning("Aviso", "Por favor, adicione pelo menos uma OS.")
        return
    if not prazo:
        messagebox.showwarning("Aviso", "Por favor, insira o prazo de entrega.")
        return
    
    os_list = list(numero_OS)
    cria_OS(prazo, os_list)

def on_closing():
    if messagebox.askokcancel("Sair", "Tem certeza que deseja fechar o programa?"):
        sys.exit()

def limpar_orcamento():
    numero_OS.clear()
    lista_valores.set('')
    N_Ordem.set('')
    prazo_entrega.set('')
    print("Lista de orçamentos limpa.")

# Inicialização da GUI
root = ttk.Window(themename="superhero")
root.geometry("800x400")
root.title("Gerador de Orçamentos")

# Variáveis
N_Ordem = StringVar()
numero_OS = []
lista_valores = StringVar()
prazo_entrega = StringVar()

root.protocol("WM_DELETE_WINDOW", on_closing)

# Widgets
label = ttk.Label(root, text="Digite o número de uma OS e clique em adicionar:", font=("Helvetica", 12))
label.pack(pady=5)

entry = ttk.Entry(root, textvariable=N_Ordem, width=50)
entry.pack(pady=3)
entry.bind("<Return>", lambda event=None: adicionar_valor())

botao_adicionar = ttk.Button(root, text="Adicionar OS", command=adicionar_valor, bootstyle="info")
botao_adicionar.pack(pady=10)

label_lista_header = ttk.Label(root, text="OSs a serem processadas:", font=("Helvetica", 10, "bold"))
label_lista_header.pack(pady=5)
label_lista = ttk.Label(root, textvariable=lista_valores, wraplength=700)
label_lista.pack(pady=5)

label = ttk.Label(root, text="Prazo de Entrega: ", font=("Helvetica", 12))
label.pack(pady=5)

entry = ttk.Entry(root, textvariable=prazo_entrega, width=50)
entry.pack(pady=5)

frame_btn = ttk.Frame(root)
frame_btn.pack(expand=False, anchor=CENTER)

botao_prosseguir = ttk.Button(frame_btn, text="Criar Orçamentos", command=prosseguir, bootstyle="info", width=15)
botao_prosseguir.pack(side=LEFT, padx=10, pady=10)

button_limpar = ttk.Button(frame_btn, text="Limpar", command=limpar_orcamento, bootstyle="warning", width=15)
button_limpar.pack(side=LEFT, padx=10, pady=10)

button_cancelar = ttk.Button(frame_btn, text="Sair", command=on_closing, bootstyle="danger", width=15)
button_cancelar.pack(side=LEFT, padx=10, pady=10)

root.mainloop()