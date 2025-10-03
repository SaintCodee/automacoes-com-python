from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from ttkbootstrap import Style
from ttkbootstrap.widgets import Entry, Button, Frame, Label
import tkinter as tk
from tkinter import Listbox, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import StringVar
from selenium.webdriver.support.ui import Select
import config # O arquivo config.py (NÃO PUBLICADO) contém credenciais e URLs

# ====================================================================
# VARIÁVEIS DE CONFIGURAÇÃO E UTILS
# ====================================================================

TIMEOUT = 20 
pecas_lista = []
quantidades_lista = []
prazo_entrega = "30"

def esperar_e_clicar(navegador, by_locator):
    WebDriverWait(navegador, TIMEOUT).until(
        EC.element_to_be_clickable(by_locator)
    ).click()

def esperar_e_enviar_teclas(navegador, by_locator, chaves):
    WebDriverWait(navegador, TIMEOUT).until(
        EC.visibility_of_element_located(by_locator)
    ).send_keys(chaves)

# ====================================================================
# FUNÇÕES DE AUTOMAÇÃO
# ====================================================================

def iniciar_automacao():
    service = Service(EdgeChromiumDriverManager().install())
    navegador = webdriver.Edge(service=service)
    
    site = config.ARKMEDS_URL
    navegador.get(site)
    
    # Simula o login. As credenciais reais são carregadas via 'config.py'.
    input("Pressione ENTER no console após fazer o login manual...") 
    return navegador

def coletar_maior_numero_os(navegador) -> int:
    def coletar_maior_os(navegador, xpath_dropdown: str, xpath_th_data: str, xpath_td_os: str) -> int | None:
        dropdown = WebDriverWait(navegador, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, xpath_dropdown))
        )
        select = Select(dropdown)
        select.select_by_visible_text("50")
        
        esperar_e_clicar(navegador, (By.XPATH, xpath_th_data))
        time.sleep(2)

        elementos_os = navegador.find_elements(By.XPATH, xpath_td_os)
        numeros_os = []
        for el in elementos_os[:50]:
            try:
                num = int(el.text.strip())
                if num < 2000:
                    numeros_os.append(num)
            except ValueError:
                continue

        return max(numeros_os) if numeros_os else None

    # Coleta OS da aba principal
    maior_os_principal = coletar_maior_os(
        navegador,
        '/html/body/div[2]/div[3]/div/div[2]/div[11]/div[1]/div/div/div/div[2]/div/div/div[1]/label/select',
        '//*[@id="tabela_orcamentos_cabecalho"]/th[11]',
        '//*[@id="table_orcamentos_aguardando"]/tbody/tr/td[7]'
    )

    # Troca para segunda aba
    esperar_e_clicar(navegador, (By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/ul/li[2]/a'))
    time.sleep(3)

    # Coleta OS da aba 2
    maior_os_aba2 = coletar_maior_os(
        navegador,
        '/html/body/div[2]/div[3]/div/div[2]/div[11]/div[2]/div/div/div/div[2]/div/div/div[1]/label/select',
        '//*[@id="tabela3"]/th[10]',
        '/html/body/div[2]/div[3]/div/div[2]/div[11]/div[2]/div/div/div/div[2]/div/div/table/tbody/tr/td[6]'
    )

    todos_os_validos = [os for os in [maior_os_principal, maior_os_aba2] if os is not None]
    numero_os_global = max(todos_os_validos) + 1 if todos_os_validos else 1

    # Número de OS anonimizado para publicação
    return 1999

def atualizar_campos_tipo_os():
    tipo = tipo_os_var.get().lower()
    frame_quant_pecas.pack_forget()
    frame_button_listbox.pack_forget()
    frame_servicos.pack_forget()
    
    if tipo == "serviços":
        frame_servicos.pack(pady=20, padx=20)
    elif tipo in ["peças", "peças e serviços"]:
        frame_quant_pecas.pack(pady=10, padx=20)
        frame_button_listbox.pack(pady=20, padx=20)

def adicionar_peca():
    peca = pecas_var.get().strip()
    quantidade_peca = pecas_qnt_var.get().strip()
    
    if not peca or not quantidade_peca:
        messagebox.showerror("Erro", "Por favor, insira a peça e a quantidade.")
        return
    
    try:
        quantidade_peca = int(quantidade_peca)
        if quantidade_peca <= 0:
             messagebox.showerror("Erro", "A quantidade deve ser um número positivo.")
             return
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número inteiro.")
        return
    
    listbox_pecas.insert(tk.END, f"{peca} - {quantidade_peca}")
    pecas_lista.append(peca)
    quantidades_lista.append(quantidade_peca)
    pecas_var.set("")
    pecas_qnt_var.set("")

def proxima_etapa():
    tipo_escolhido = tipo_os_var.get().strip()
    if not tipo_escolhido:
        messagebox.showerror("Erro", "Por favor, selecione o Tipo de OS.")
        return
    
    navegador = iniciar_automacao()
    action = ActionChains(navegador)

    try:
        numero_os_global = coletar_maior_numero_os(navegador)

        #Primeira etapa, seleção do cliente
        esperar_e_clicar(navegador, (By.ID,'adicionar-orcamento'))
        esperar_e_clicar(navegador, (By.XPATH,'//*[@id="teste-modal"]/div[2]/span')) 
        
        # Seleção do tipo de orçamento (Serviços ou Peças)
        if tipo_escolhido.lower() == "serviços":
            opcoes = WebDriverWait(navegador, 10).until(
                EC.visibility_of_all_elements_located((By.CSS_SELECTOR, '.select2-results__option'))
            )
            if len(opcoes) >= 4:
                opcoes[3].click()
            else:
                messagebox.showerror("Erro", "Não foi possível encontrar a opção 'Serviços'.")
                return
        else:
            action.send_keys(tipo_escolhido).send_keys(Keys.ENTER).perform()
        
        time.sleep(1)
        esperar_e_clicar(navegador, (By.XPATH, '//*[@id="proxima-etapa-orcamento"]'))

        # --- Lógica de preenchimento para Serviços e Peças ---

        if tipo_os_var.get() == "Serviços":
            esperar_e_clicar(navegador, (By.XPATH, '//*[@id="select2-orcamento-solicitante-servico-container"]'))
            observacoes = observacoes_text.get("1.0",tk.END).strip()
            solicitante = solicitante_var.get().strip()
            action.send_keys(solicitante).send_keys(Keys.ENTER).perform()
            time.sleep(1)
            
            # Segunda etapa
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="comeco_orcamento_servicos"]/div/div[2]/div/div[3]/button[2]'))
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="form-servicos"]/div/div[1]/div/div/div/span[2]/span[1]/span/span[2]/b'))
            action.send_keys(tipo_servico_var.get()).perform()
            time.sleep(5)
            action.send_keys(Keys.ENTER).perform()
            esperar_e_clicar(navegador, (By.ID,'quantidade-servico'))
            action.send_keys("1").perform()
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="botao-servico"]/strong'))
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="comeco_orcamento_servicos"]/div/div[2]/div/div[3]/button[2]/i'))
            
            # Terceira etapa
            esperar_e_clicar(navegador, (By.ID,'valor_desconto'))
            action.send_keys(Keys.TAB).send_keys("Faturado").send_keys(Keys.ENTER).perform()
            action.send_keys(Keys.TAB).send_keys("Outros").send_keys(Keys.ENTER).perform()
            action.send_keys(Keys.TAB).send_keys("Boleto 20 dias").perform()
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="comeco_orcamento_servicos"]/div/div[2]/div/div[3]/button[2]/i'))
            
            # Quarta Etapa - Prazo de Entrega
            os_input = navegador.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/div[5]/div/div[2]/div/div[2]/div[4]/form/div[2]/div[3]/div/div/input')
            os_input.clear()
            os_input.send_keys(str(numero_os_global))
            esperar_e_clicar(navegador, (By.ID,'prazo_entrega'))
            action.send_keys(prazo_entrega).perform()
            esperar_e_clicar(navegador, (By.ID,'validade'))
            action.send_keys("90").perform()
            
            # Remoção do nome real (Ícaro)
            action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB)
            action.send_keys(
                (f"Equipamento: {equipamento_var.get().strip()}, " if equipamento_var.get().strip() else "") +
                (f"Modelo: {modelo_var.get().strip()}, " if modelo_var.get().strip() else "") +
                (f"Número de Série: {serie_var.get().strip()}" if serie_var.get().strip() else "") +
                (f"\n\nServiços a serem executados:\n {observacoes}" if observacoes else "")).perform()
            action.send_keys("\n\nObs: A garantia não cobre peças não substituídas, mau uso ou serviços que não foram realizados.").perform()
            
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="form4"]/div[2]/div[8]/div[1]/div/span/span[1]/span/span[2]'))
            action.send_keys("Nome Anonimizado").send_keys(Keys.ENTER).perform() # Substituição do nome
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="form4"]/div[2]/div[8]/div[2]/div/div'))


        elif tipo_os_var.get() == "Peças":
            
            esperar_e_clicar(navegador, (By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/div[5]/div/div[2]/div/div[2]/div[1]/form/span/span[1]/span/span[1]'))
            observacoes = observacoes_text.get("1.0",tk.END).strip()

            # Segunda Etapa - Seleção do Solicitante
            solicitante = solicitante_var.get().strip()
            action.send_keys(solicitante).send_keys(Keys.ENTER).perform()
            
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="main-content-wrapper"]/div/div[2]/div[5]/div/div[2]/div/div[2]/div[5]/button[2]'))
            
            # Segunda Etapa - Seleção das Peças
            navegador.maximize_window()
            for i, elemento in enumerate(pecas_lista):
                esperar_e_clicar(navegador, (By.XPATH,'//*[@id="select2-orcamento-pecas-container"]'))
                action.send_keys(pecas_lista[i]).send_keys(Keys.ENTER).perform()
                action.send_keys(Keys.TAB).send_keys(quantidades_lista[i]).perform()
                action.send_keys(Keys.TAB, Keys.UP, Keys.TAB).send_keys("90").perform() # Prazo de entrega
                esperar_e_clicar(navegador, (By.XPATH,'//*[@id="botao-salvar-peca"]/strong'))
            
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="main-content-wrapper"]/div/div[2]/div[5]/div/div[2]/div/div[2]/div[5]/button[2]'))
            
            # Tipo de pagamento
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="main-content-wrapper"]/div/div[2]/div[5]/div/div[2]/div/div[2]/div[5]/button[2]'))
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="form3"]/div[4]/div/div/span/span[1]/span'))
            action.send_keys(Keys.DOWN, Keys.DOWN, Keys.DOWN, Keys.DOWN, Keys.ENTER).perform()
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="form3"]/div[5]/div/div/span/span[1]/span'))
            action.send_keys(Keys.DOWN, Keys.DOWN, Keys.DOWN, Keys.ENTER).perform()
            esperar_e_enviar_teclas(navegador, (By.ID,'outros'), "Boleto 20 dias")
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="main-content-wrapper"]/div/div[2]/div[5]/div/div[2]/div/div[2]/div[5]/button[2]'))

            # Quarta Etapa - Prazo de Entrega
            os_input = navegador.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/div[5]/div/div[2]/div/div[2]/div[4]/form/div[2]/div[3]/div/div/input')
            os_input.clear()
            os_input.send_keys(str(numero_os_global))
            esperar_e_clicar(navegador, (By.ID,'prazo_entrega'))
            action.send_keys(prazo_entrega).perform()
            esperar_e_clicar(navegador, (By.ID,'validade'))
            action.send_keys("90").perform()
            
            action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB)
            action.send_keys(
                (f"Equipamento: {equipamento_var.get().strip()}, " if equipamento_var.get().strip() else "") +
                (f"Modelo: {modelo_var.get().strip()}, " if modelo_var.get().strip() else "") +
                (f"Número de Série: {serie_var.get().strip()}" if serie_var.get().strip() else "") +
                (f"\n\nServiços a serem executados:\n {observacoes}" if observacoes else "")).perform()
            action.send_keys("\n\nObs: A garantia não cobre peças não substituídas, mau uso ou serviços que não foram realizados.").perform()
            
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="form4"]/div[2]/div[8]/div[1]/div/span/span[1]/span/span[2]'))
            action.send_keys("Nome Anonimizado").send_keys(Keys.ENTER).perform() # Substituição do nome
            esperar_e_clicar(navegador, (By.XPATH,'//*[@id="form4"]/div[2]/div[8]/div[2]/div/div'))


    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    finally:
        input('Fim do Programa\n')

# ====================================================================
# CÓDIGO DA INTERFACE GRÁFICA (MANTIDO IDÊNTICO)
# ====================================================================

# Inicializa as listas de peças e quantidades
# (As listas globais já estão no topo)

# Configuração da janela principal
app = tk.Tk()
style = Style(theme='superhero')
app.title("Orçamentos Sem OS")
app.geometry('1100x700')

# Variáveis para armazenar os dados dos campos de entrada
pecas_var = tk.StringVar()
pecas_qnt_var = tk.StringVar()
solicitante_var = tk.StringVar()
equipamento_var = tk.StringVar()
modelo_var = tk.StringVar()
serie_var = tk.StringVar()
tipo_os_var = tk.StringVar()


frame = Frame(app)
frame.pack(pady=20)

tipo_os_frame = Frame(app)
tipo_os_frame.pack(pady=10, padx=20, anchor='center')

opcoes_tipo_os = [
    ("Peças", "Peças"),
    ("Serviços", "Serviços"),
]

ttk.Label(tipo_os_frame, text="Tipo de OS:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
for i, (texto, valor) in enumerate(opcoes_tipo_os):
    ttk.Radiobutton(
        tipo_os_frame,
        text=texto,
        variable=tipo_os_var,
        value=valor,
        bootstyle="info",
        command=atualizar_campos_tipo_os
    ).grid(row=0, column=i+1, padx=5, pady=5, sticky='w')


# Frame para o solicitante
solicitante_frame = Frame(app)
solicitante_frame.pack(pady=20, padx=20)

ttk.Label(solicitante_frame, text="Cliente").grid(row=0, column=0, padx=5, pady=5)
ttk.Entry(solicitante_frame, textvariable=solicitante_var, width=30).grid(row=0, column=1, padx=5, pady=5)

frame_labels_entrys = Frame(app)
frame_labels_entrys.pack(pady=10, padx=20)

labels = ["Equipamento", "Modelo", "Número de Série"]
for i, label_text in enumerate(labels):
    ttk.Label(frame_labels_entrys, text=label_text).grid(row=0, column=i, padx=5, pady=6)

entrys = [equipamento_var, modelo_var, serie_var]
for i, entry in enumerate(entrys):
      ttk.Entry(frame_labels_entrys, textvariable=entry, width=30).grid(row=1, column=i, padx=5, pady=6)

frame_quant_pecas = Frame(app)
frame_quant_pecas.pack(pady=10, padx=20)

# Label e frame para as peças e quantidades
Label(frame_quant_pecas, text="Digite o descritivo e a quantidade das peças:").pack(pady=10)
frame_pecas = Frame(frame)
frame_pecas.pack(expand=FALSE, anchor=CENTER)

# Campo de entrada para o descritivo da peça
entry_pecas = Entry(frame_quant_pecas, textvariable=pecas_var, width=30)
entry_pecas.pack(side=LEFT, padx=10, pady=10)

# Campo de entrada para a quantidade da peça
entry_pecas_qnt = ttk.Entry(frame_quant_pecas, textvariable=pecas_qnt_var, width=5)
entry_pecas_qnt.pack(side=LEFT, padx=10, pady=10)

frame_button_listbox = Frame(app)
frame_button_listbox.pack(pady=20, padx=20)
# Botão para adicionar a peça à lista de peças
button_add_peca = ttk.Button(frame_button_listbox, text="Adicionar Peça", bootstyle="superhero", command=adicionar_peca, width=20)
button_add_peca.grid(row=2, column=0, columnspan=5, pady=10)

# Listbox para mostrar as peças adicionadas
listbox_pecas = Listbox(frame_button_listbox, width=40)
listbox_pecas.grid(row=3, column=0, columnspan=5, pady=10)

# Botão para prosseguir para a próxima etapa
button_prosseguir = Button(frame_button_listbox, text="Criar Orçamento", command=proxima_etapa, style='success')
button_prosseguir.grid(row=4, column=0, columnspan=5, pady=20)

# Frame para campos de Serviços
frame_servicos = Frame(app)

# Campo para o tipo de serviço
ttk.Label(frame_servicos, text="Tipo de Serviço:").pack(anchor='w', padx=5, pady=5)
tipo_servico_var = tk.StringVar()
ttk.Entry(frame_servicos, textvariable=tipo_servico_var, width=40).pack(fill='x', padx=5, pady=5)

# Campo para observações (caixa de texto grande)
ttk.Label(frame_servicos, text="Observações:").pack(anchor='w', padx=5, pady=5)
observacoes_text = tk.Text(frame_servicos, width=60, height=6)
observacoes_text.pack(fill='x', padx=5, pady=5)

# Botão para prosseguir para a próxima etapa (Serviços)
ttk.Button(frame_servicos,text="Criar Orçamento",command=proxima_etapa,bootstyle="success",width=20).pack(pady=20)

atualizar_campos_tipo_os()

app.mainloop()