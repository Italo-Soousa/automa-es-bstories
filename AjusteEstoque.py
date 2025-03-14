from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import tkinter as tk
import sys
import time

class RedirectText:
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, message):
        self.output.insert(tk.END, message)  # Insere a mensagem no final do terminal
        self.output.see(tk.END)  # Auto-scroll para a última linha

    def flush(self):
        pass  # Necessário para evitar erros com sys.stdout (stdout requer um método flush)

def ajusteEstoque(produto, qtd_subtraida, justificativa,  empresa, opperacao):
    def criar_navegador():
        caminho_drive = "/Users/italosousa/Documents/Relatórios Fornecedores/Relatórios LD"

        chrome_options = Options()
        prefs = {
                "download.default_directory": caminho_drive,  # Define a pasta do Google Drive
                "download.prompt_for_download": False,  # Não perguntar onde salvar
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "plugins.always_open_pdf_externally": True,  # Evita que PDFs sejam abertos no navegador
                "profile.default_content_settings.popups": 0,
                "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
                "download.extensions_to_open": "applications/vnd.ms-excel"  # Força arquivos Excel a baixarem corretamente
            }

        chrome_options.add_experimental_option("prefs", prefs)

        navegador = webdriver.Chrome()
        espera = WebDriverWait(navegador, 20)
        return navegador, espera

    def fazer_login(navegador, espera, login, senha):
        """
        Realiza o login no sistema.
        """
        navegador.get("https://erp.microvix.com.br/")
        espera.until(EC.presence_of_element_located((By.ID, "f_login"))).send_keys(login)
        navegador.find_element(By.ID, "f_senha").send_keys(senha)
        navegador.find_element(By.XPATH, '//*[@id="form_login"]/button[1]').click()

    def selecionar_empresa(navegador, espera, empresa_index):
        """
        Seleciona a empresa desejada no dropdown.
        """
        botao_dropdown = espera.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn.dropdown-toggle.bs-placeholder.btn-primary")
        ))
        botao_dropdown.click()
        xpath_empresa = f'//*[@id="frm_selecao_empresa_login"]/div/div/div/div[2]/ul/li[{empresa_index}]/a'
        espera.until(EC.element_to_be_clickable((By.XPATH, xpath_empresa))).click()
        espera.until(EC.element_to_be_clickable((By.ID, "btnselecionar_empresa"))).click()

        time.sleep(2)
        espera.until(EC.element_to_be_clickable((By.ID, 'frente-logo-hamburger'))).click()

    def fazerAjuste(navegador, espera, produto, qtd_subtraida):
        # Indo até os produtos
        espera.until(EC.element_to_be_clickable((By.ID, 'liModulo_8'))).click()
        espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="liModulo_8"]/ul/li[1]/a'))).click()
        navegador.find_element(By.XPATH, '//*[@id="liModulo_8"]/ul/li[1]/ul/li[2]/a').click()

        # Selecionando o produto
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)
        espera.until(EC.element_to_be_clickable((By.ID, 'inputPesquisaProdutos'))).send_keys(produto)
        espera.until(EC.element_to_be_clickable((By.ID, 'pesquisarProdutos'))).click()

        # Indo até o ajuste
        espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainAppBody"]/div/div[3]/div/div[2]/div/div/div[1]/table/tbody/tr[1]/td[1]'))).click()
        menu_lateral = espera.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'v-sidebar-menu vsm_expanded')]")))
        navegador.execute_script("arguments[0].style.visibility='visible'; arguments[0].style.display='block';", menu_lateral)
        botao_saldo = espera.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Saldo')]")))
        botao_saldo.click()
        botao_ajuste = espera.until(EC.element_to_be_clickable((By.XPATH, "//a[.//span[contains(text(), 'Ajustar saldo')]]")))
        botao_ajuste.click()

        # Logica para o calculo do ajuste
        time.sleep(4)
        numero = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[3]/td[3]/font/b')
        valor = numero.text.strip()  # Remove espaços extras

        # Converte "7,00" para 7.00 (substitui a vírgula por ponto e converte para float)
        valor = float(valor.replace(',', '.'))

        num_subtraido = float(qtd_subtraida)

        # Agora a subtração funcionará corretamente

        if opperacao == "-":
            ajuste = num_subtraido - valor
        elif opperacao == "+":
            ajuste = num_subtraido + valor

        # Formata o número para o formato correto: duas casas decimais e vírgula
        ajuste_formatado = f"{ajuste:.2f}".replace('.', ',')

        # Insere o valor formatado no campo de input
        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[4]/td[3]/input').clear()
        navegador.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[4]/td[3]/input').send_keys(ajuste_formatado)

        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[5]/td[3]/font/input').send_keys(justificativa)

        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/form/table/tbody/tr[6]/td[2]/font/input').click()

        mensagem = print(f'Foi feito o ajusto de estoque do produto: {produto}, que tinha um saldo de: {valor}, onde foi subtraido {qtd_subtraida}, ficando com um saldo de {ajuste}')

        time.sleep(5)

        return mensagem


    navegador, espera = criar_navegador()

    try:

        fazer_login(navegador, espera, "italo.meca", "@86Kizzmacca6")

        selecionar_empresa(navegador, espera, empresa)

        fazerAjuste(navegador, espera, produto, qtd_subtraida)
    finally:
        navegador.quit()

    time.sleep(5)

# Criando a interface principal
root = tk.Tk()
root.title('Automações B.stories')
root.geometry("411x466")  # Ajustando tamanho para caber o terminal

# Criando os Frames
top_frame = tk.Frame(root, height=50, bg="white")
top_frame.pack(side="top", fill="x")

left_frame = tk.Frame(root, width=200)
left_frame.pack(side="left", fill="y")

separator = tk.Frame(root, width=2, bg="white")
separator.pack(side="left", fill="y")

content_area = tk.Frame(root)
content_area.pack(side="left", expand=True, fill="both")

# Terminal area
terminal_frame = tk.Frame(root, width=300)
terminal_frame.pack(side="right", fill="both", expand=True)

# Adicionando título ao top_frame
titulo = tk.Label(top_frame, text='Automações B.stories', font='Helvetica 16 bold', fg="black", bg="white")
titulo.pack(pady=10)

# Botões na barra lateral (left_frame)
geracaoRelatorios = tk.Button(left_frame, text='Geração Relatórios')
geracaoRelatorios.pack(pady=10, padx=15, fill="x")

ajuste = tk.Button(left_frame, text='Ajuste Estoque', fg='blue')
ajuste.pack(pady=10, padx=15, fill="x")

subirEstoque = tk.Button(left_frame, text='Subir Estoque')
subirEstoque.pack(pady=10, padx=15, fill="x")

tituloAjusteEstoque = tk.Label(content_area, text= 'Ajuste de Estoque', font='Helvetica 16 bold')
tituloAjusteEstoque.grid(row= 0, column= 1, sticky="ew")

# Qual produto
referenciaProduto = tk.Label(content_area, text='Qual o código do produto: ')
referenciaProduto.grid(row=1, column=1, padx=10, pady=5)

entryProduto = tk.Entry(content_area)
entryProduto.grid(row=2, column=1, padx=10, pady=5)

# Quantidade retirada
qtdRetirada = tk.Label(content_area, text='Qual a quantidade retirada: ')
qtdRetirada.grid(row=3, column=1, padx=10, pady=5)

entryRetirada = tk.Entry(content_area)
entryRetirada.grid(row=4, column=1, padx=10, pady=5)

# Justificativa
JustRetirada = tk.Label(content_area, text='Qual a justificativa do ajuste: ')
JustRetirada.grid(row=5, column=1, padx=10, pady=5)

entryJustificativa = tk.Entry(content_area)
entryJustificativa.grid(row=6, column=1, padx=10, pady=5)


# Criando um frame exclusivo para os botões
iframe_botoes = tk.Frame(content_area)
iframe_botoes.grid(row=8, column=1, columnspan=3, pady=10)

# Botões dentro do novo Frame
botao_LD = tk.Button(iframe_botoes, text='LD', command= lambda: ajusteEstoque(entryProduto.get(), entryRetirada.get(), entryJustificativa.get(),  2, "-"))
botao_LD.pack(side=tk.LEFT, padx=3)

botao_MARY = tk.Button(iframe_botoes, text='MARY', command= lambda: ajusteEstoque(entryProduto.get(), entryRetirada.get(), entryJustificativa.get(), 3,  "-"))
botao_MARY.pack(side=tk.LEFT, padx=3)
#
botao_MECA = tk.Button(iframe_botoes, text='MECA', command= lambda: ajusteEstoque(entryProduto.get(), entryRetirada.get(), entryJustificativa.get(), 1,  "-"))
botao_MECA.pack(side=tk.LEFT, padx=3)

# Criando a área do "terminal" menor na interface
terminal_frame = tk.Frame(content_area)
terminal_frame.grid(row=9, column=1, columnspan=3, pady=5, sticky="nsew")

terminal_label = tk.Label(terminal_frame, text="Terminal", font="Helvetica 10 bold")
terminal_label.pack(anchor="w", padx=5, pady=2)

# Criando um Frame extra branco para a borda do terminal
terminal_border = tk.Frame(terminal_frame, bg="white")  # **Borda Branca**
terminal_border.pack(padx=5, pady=5)

# **Terminal Menor dentro do Frame com borda branca**
terminal_text = tk.Text(terminal_border, bg="black", fg="white", insertbackground="white", wrap="word",
                        height=5, width=32, font=("Courier", 9))

terminal_text.pack(expand=False, fill="both", padx=5, pady=2)

# Redirecionando stdout para o widget Text
sys.stdout = RedirectText(terminal_text)


root.mainloop()