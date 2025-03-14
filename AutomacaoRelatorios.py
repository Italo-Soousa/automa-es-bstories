import tkinter as tk
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from tkinter import messagebox
import os
import time
import sys

class RedirectText:
    """ Classe para redirecionar a saída do terminal para um widget Text """
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)  # Rolagem automática para a última linha

    def flush(self):
        pass  # Necessário para compatibilidade com sys.stdout

def validar_entrada(funcao, data_inicial, data_final):
    if not data_inicial.strip() or not data_final.strip():
        messagebox.showwarning("Aviso", "Por favor, preencha todas as datas antes de continuar.")
    else:
        funcao(data_inicial, data_final)

def relatoriosLD(data_inicio, data_fim):
    def criar_navegador():
        """
        Inicializa o navegador e configura as opções para baixar arquivos no Google Drive.
        """
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

        navegador = webdriver.Chrome(options=chrome_options)
        navegador.maximize_window()
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

    def acessar_relatorio(navegador, espera):
        """
        Navega até o relatório de produtos e serviços vendidos.
        """
        espera.until(EC.element_to_be_clickable((By.ID, 'liModulo_10'))).click()
        relatorios = navegador.find_element(By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/a')
        navegador.execute_script('arguments[0].scrollIntoView(true)', relatorios)
        relatorios.click()

        # Rolando até o relatório desejado
        navegador.execute_script(
            "arguments[0].scrollIntoView(true);",
            navegador.find_element(By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/ul/li[15]/a')
        )
        time.sleep(5)
        espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/ul/li[18]/a'))).click()

    def configurar_relatorio(navegador, espera, data_inicial, data_final):
        """
        Configura o filtro do relatório com as datas e outras opções.
        """
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        for campo, valor in [("f_data1", data_inicial), ("f_data2", data_final)]:
            campo_data = espera.until(EC.element_to_be_clickable((By.ID, campo)))
            campo_data.click()
            campo_data.send_keys(Keys.CONTROL + "a")
            campo_data.send_keys(Keys.BACKSPACE)
            campo_data.send_keys(valor)
            time.sleep(1)

            # Seleciona e desseleciona empresas
            navegador.find_element(By.ID, 'empresas_3').click()
            navegador.find_element(By.ID, 'empresas_3').click()

    def selecionar_fornecedor(navegador, espera, fornecedor, id_fornecedor):

        """
        Seleciona o fornecedor na janela de busca.
        """
        janelas_antes = navegador.window_handles
        lupa_fornecedor = navegador.find_element(By.XPATH, '//*[@id="clientefornecedor_btpesquisar"]/td[2]')
        lupa_fornecedor.click()
        WebDriverWait(navegador, 10).until(
            lambda d: len(d.window_handles) > len(janelas_antes)
        )

        nova_janela = [j for j in navegador.window_handles if j not in janelas_antes][0]
        navegador.switch_to.window(nova_janela)
        navegador.maximize_window()

        # Buscar e selecionar fornecedor
        espera.until(EC.element_to_be_clickable((By.ID, 'pesquisa'))).send_keys(fornecedor)
        navegador.find_element(By.ID, 'sub').click()
        espera.until(EC.element_to_be_clickable((By.ID, f'chk_{id_fornecedor}'))).click()
        navegador.find_element(By.ID, 'btFPAdicionar').click()

        navegador.switch_to.window(janelas_antes[0])

    def selecionar_marca(navegador, espera, marca, id_marca):
        """
        Seleciona a marca na janela de busca.
        """
        # Entra no iframe
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        # Controle de janelas
        janelas_antes = navegador.window_handles

        # Localizar e clicar na lupa de marcas
        lupa_marca = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="marcas_btpesquisar"]/td[1]')))
        navegador.execute_script("arguments[0].scrollIntoView(true);", lupa_marca)
        lupa_marca.click()

        # Controle de janelas
        WebDriverWait(navegador, 10).until(
            lambda d: len(d.window_handles) > len(janelas_antes)
        )

        # Alternar para a nova janela
        nova_janela = [j for j in navegador.window_handles if j not in janelas_antes][0]
        navegador.switch_to.window(nova_janela)
        navegador.maximize_window()

        # Realizar a pesquisa e selecionar a marca
        espera.until(EC.element_to_be_clickable((By.ID, 'pesquisa'))).send_keys(marca)
        navegador.find_element(By.ID, 'sub').click()
        espera.until(EC.element_to_be_clickable((By.ID, f'chk_{id_marca}'))).click()
        navegador.find_element(By.ID, 'btFPAdicionar').click()

        # Voltar para a janela original
        navegador.switch_to.window(janelas_antes[0])
        print(f"Marca '{marca}' selecionada com sucesso!")

    def finalizar_relatorio(navegador, espera):

        """
        Finaliza a configuração do relatório e exporta para Excel.
        """
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        # Seleciona tipo de custo
        tipo_custo = Select(navegador.find_element(By.ID, 'tipo_custo'))
        tipo_custo.select_by_index(5)

        # Ajustes no relatório
        time.sleep(1)
        navegador.find_element(By.XPATH,
                               '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[28]/td[2]/input[2]').click()
        navegador.find_element(By.XPATH,
                               '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[30]/td[2]/input[3]').click()
        navegador.find_element(By.XPATH, '//*[@id="mostralistar"]/td[2]/input[2]').click()
        navegador.find_element(By.ID, 'exibir_barras').click()

        # Gerar o relatório
        navegador.find_element(By.CLASS_NAME, 'btn-submit').click()
        espera.until(EC.element_to_be_clickable((By.ID, 'botaoExportarXLS'))).click()

    def acessar_relatorio_novamente(navegador, espera, data_inicial, data_final):
        """
        Sai do iframe, volta ao menu de relatórios e acessa o relatório desejado.
        """
        # Sair do iframe e voltar ao conteúdo principal
        navegador.switch_to.default_content()

        # Clica no relatório desejado
        relatorio = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/ul/li[18]/a')))
        relatorio.click()

        # Voltar para o iframe após acessar o relatório
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        for campo, valor in [("f_data1", data_inicial), ("f_data2", data_final)]:
            campo_data = espera.until(EC.element_to_be_clickable((By.ID, campo)))
            campo_data.click()
            campo_data.send_keys(Keys.CONTROL + "a")
            campo_data.send_keys(Keys.BACKSPACE)
            campo_data.send_keys(valor)

        # Seleciona e desseleciona empresas
        navegador.find_element(By.ID, 'empresas_3').click()
        navegador.find_element(By.ID, 'empresas_3').click()

    def relatorio_lvmh(navegador, espera, data_inicial, data_final, fornecedor, cod_fornecedor, marca, cod_marca):

        print(f'Gerando relatório para fornecedor: LVMH (ID: 248), marca: DIOR (ID: 6)')

        # Voltar ao menu de relatórios e acessar novamente
        acessar_relatorio_novamente(navegador, espera, data_inicial, data_final)

        # Configurar o relatório para o fornecedor atual
        selecionar_fornecedor(navegador, espera, fornecedor, cod_fornecedor)

        time.sleep(2)

        # Seleciona a marca
        selecionar_marca(navegador, espera, marca, cod_marca)

        # Finalizar e salvar o relatório
        finalizar_relatorio(navegador, espera)
        time.sleep(3)  # Evitar sobreposição entre relatórios

        print('Relatório Lvmh gerado com sucesso!')

    def relatorio_shiseido(navegador, espera, data_inicial, data_final, fornecedor, cod_fornecedor):

        print(f'Gerando relatório para fornecedor: SHISEIDO (ID: 357)')

        # Voltar ao menu de relatórios e acessar novamente
        acessar_relatorio_novamente(navegador, espera, data_inicial, data_final)

        # Configurar o relatório para o fornecedor atual
        selecionar_fornecedor(navegador, espera, fornecedor, cod_fornecedor)
        time.sleep(2)

        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        # Seleciona tipo de custo
        tipo_custo = Select(navegador.find_element(By.ID, 'tipo_custo'))
        tipo_custo.select_by_index(5)

        # Ajustes no relatório
        navegador.find_element(By.XPATH,
                               '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[28]/td[2]/input[2]').click()
        navegador.find_element(By.XPATH, '//*[@id="mostralistar"]/td[2]/input[2]').click()
        navegador.find_element(By.ID, 'exibir_barras').click()

        # Gerar o relatório
        navegador.find_element(By.CLASS_NAME, 'btn-submit').click()
        espera.until(EC.element_to_be_clickable((By.ID, 'botaoExportarXLS'))).click()

        time.sleep(3)

        print('Relatório Shiseido gerado com sucesso!')

    def renomear_relatorios(pasta):
        """
        Renomeia os relatórios baixados e convertidos conforme o mapeamento especificado.
        """
        # Definir o mapeamento de nomes corrigindo possíveis espaços extras
        mapeamento = {
            "ProdutosServicosVendidos.xls": "PUIG-LD.xls",
            "ProdutosServicosVendidos (1).xls": "DBB-LD.xls",
            "ProdutosServicosVendidos (2).xls": "LOREAL-LD.xls",
            "ProdutosServicosVendidos (3).xls": "EXCELLENCE-LD.xls",
            "ProdutosServicosVendidos (4).xls": "PRESTIGE-LD.xls",
            "ProdutosServicosVendidos (5).xls": "TABATINGA-LD.xls",
            "ProdutosServicosVendidos (6).xls": "JUNO-LD.xls",
            "ProdutosServicosVendidos (7).xls": "COTY-LD.xls",
            "ProdutosServicosVendidos (8).xls": "SHISEIDO-LD.xls",
            "ProdutosServicosVendidos (9).xls": "LVMH-LD.xls"
        }

        # Confirma se a pasta existe
        if not os.path.exists(pasta):
            print(f"Erro: A pasta '{pasta}' não existe!")
            return

        # Lista os arquivos exatos na pasta
        arquivos = os.listdir(pasta)
        print(f"\n📂 Arquivos encontrados na pasta '{pasta}':")
        for arquivo in arquivos:
            print(f" - {arquivo}")

        # Renomear os arquivos
        for arquivo in arquivos:
            caminho_antigo = os.path.join(pasta, arquivo)

            # Verifica se o nome do arquivo está no mapeamento
            if arquivo in mapeamento:
                novo_nome = mapeamento[arquivo]
                caminho_novo = os.path.join(pasta, novo_nome)

                os.rename(caminho_antigo, caminho_novo)
                print(f"✅ Renomeado: {arquivo} → {novo_nome}")
            else:
                print(f"⚠ Arquivo ignorado: {arquivo} (não mapeado)")

    # Definir a pasta correta onde os arquivos foram baixados
    pasta_relatorios = "/Users/italosousa/Documents/Relatórios Fornecedores/Relatórios LD"

    def corrigir_extensao_arquivos(pasta):
        """
        Renomeia arquivos que foram salvos como .html para .xls,
        verificando se há extensões incorretas.
        """
        print("\n🔍 Verificando arquivos na pasta antes da correção:")
        arquivos = os.listdir(pasta)

        if not arquivos:
            print("⚠ Nenhum arquivo encontrado na pasta.")
            return

        for arquivo in arquivos:
            caminho_antigo = os.path.join(pasta, arquivo)

            # Verifica se o arquivo é HTML disfarçado de XLS
            if arquivo.endswith(".html") or arquivo.endswith(".xls.html"):
                caminho_novo = os.path.join(pasta, arquivo.replace(".html", "").replace(".xls", "") + ".xls")
                os.rename(caminho_antigo, caminho_novo)
                print(f"✅ Arquivo corrigido: {arquivo} → {os.path.basename(caminho_novo)}")

            # Se for um arquivo XLS já válido, apenas exibe no log
            elif arquivo.endswith(".xls"):
                print(f"✔ Arquivo já está correto: {arquivo}")

        print("\n✅ Todos os arquivos foram verificados e corrigidos!\n")

    # ---------------- Fluxo Principal ----------------
    if __name__ == "__main__":
        # Lista de fornecedores (nome e ID)
        fornecedores = [
            {"nome": "DBB", "id": 127},
            {"nome": "loreal", "id": 224},
            {"nome": "EXCELLENCE", "id": 156},
            {"nome": "PRESTIGE", "id": 313},
            {"nome": "TABATINGA", "id": 373},
            {"nome": "JUNO", "id": 213},
            {"nome": "COTY", "id": 101494}
        ]

        navegador, espera = criar_navegador()

        try:
            # Credenciais
            fazer_login(navegador, espera, "italo.meca", "@86Kizzmacca6")
            selecionar_empresa(navegador, espera, 2)
            acessar_relatorio(navegador, espera)
            configurar_relatorio(navegador, espera, data_inicio, data_fim)
            selecionar_fornecedor(navegador, espera, "puig", 318)
            finalizar_relatorio(navegador, espera)
            time.sleep(3)

            for fornecedor in fornecedores:
                print(f"Gerando relatório para fornecedor: {fornecedor['nome']} (ID: {fornecedor['id']})")

                # Voltar ao menu de relatórios e acessar novamente
                acessar_relatorio_novamente(navegador, espera, data_inicio, data_fim)

                # Configurar o relatório para o fornecedor atual
                selecionar_fornecedor(navegador, espera, fornecedor["nome"], fornecedor["id"])

                # Finalizar e salvar o relatório
                finalizar_relatorio(navegador, espera)
                time.sleep(3)  # Evitar sobreposição entre relatórios

                print(f"Relatório para {fornecedor['nome']} gerado com sucesso!")

            # Relatório shiseido que é agrupado por vendedor
            relatorio_shiseido(navegador, espera, data_inicio, data_fim, "shiseido", 357)

            # Gerando relatório LVMH que tem que selecionar marca
            relatorio_lvmh(navegador, espera, data_inicio, data_fim, "lvmh", 248, "dior", 6)

            corrigir_extensao_arquivos(pasta_relatorios)

            # **Renomear os arquivos após a conversão**
            renomear_relatorios(pasta_relatorios)

            print("Todos os relatórios foram gerados, convertidos, renomeados e tratados com sucesso.")
        finally:
            navegador.quit()

def relatorioMARY(data_inicio, data_fim):
    def criar_navegador():
        """
        Inicializa o navegador e configura as opções para baixar arquivos no Google Drive.
        """
        caminho_drive = "/Users/italosousa/Documents/Relatórios Fornecedores/Relatórios MARY"

        chrome_options = Options()
        prefs = {
            "download.default_directory": caminho_drive,  # Define a pasta do Google Drive
            "download.prompt_for_download": False,  # Não perguntar onde salvar
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        navegador = webdriver.Chrome(options=chrome_options)
        navegador.maximize_window()
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

    def acessar_relatorio(navegador, espera):
        """
        Navega até o relatório de produtos e serviços vendidos.
        """
        espera.until(EC.element_to_be_clickable((By.ID, 'liModulo_10'))).click()
        relatorios = navegador.find_element(By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/a')
        navegador.execute_script('arguments[0].scrollIntoView(true)', relatorios)
        relatorios.click()

        # Rolando até o relatório desejado
        navegador.execute_script(
            "arguments[0].scrollIntoView(true);",
            navegador.find_element(By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/ul/li[15]/a')
        )
        time.sleep(5)
        espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/ul/li[18]/a'))).click()

    def configurar_relatorio(navegador, espera, data_inicial, data_final):
        """
        Configura o filtro do relatório com as datas e outras opções.
        """
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        for campo, valor in [("f_data1", data_inicial), ("f_data2", data_final)]:
            campo_data = espera.until(EC.element_to_be_clickable((By.ID, campo)))
            campo_data.click()
            campo_data.send_keys(Keys.CONTROL + "a")
            campo_data.send_keys(Keys.BACKSPACE)
            campo_data.send_keys(valor)

        # Seleciona e desseleciona empresas
        navegador.find_element(By.ID, 'empresas_3').click()
        navegador.find_element(By.ID, 'empresas_2').click()

    def selecionar_fornecedor(navegador, espera, fornecedor, id_fornecedor):

        """
        Seleciona o fornecedor na janela de busca.
        """
        janelas_antes = navegador.window_handles
        lupa_fornecedor = navegador.find_element(By.XPATH, '//*[@id="clientefornecedor_btpesquisar"]/td[2]')
        lupa_fornecedor.click()
        WebDriverWait(navegador, 10).until(
            lambda d: len(d.window_handles) > len(janelas_antes)
        )

        nova_janela = [j for j in navegador.window_handles if j not in janelas_antes][0]
        navegador.switch_to.window(nova_janela)
        navegador.maximize_window()

        # Buscar e selecionar fornecedor
        espera.until(EC.element_to_be_clickable((By.ID, 'pesquisa'))).send_keys(fornecedor)
        navegador.find_element(By.ID, 'sub').click()
        espera.until(EC.element_to_be_clickable((By.ID, f'chk_{id_fornecedor}'))).click()
        navegador.find_element(By.ID, 'btFPAdicionar').click()

        navegador.switch_to.window(janelas_antes[0])

    def selecionar_marca(navegador, espera, marca, id_marca):
        """
        Seleciona a marca na janela de busca.
        """
        # Entra no iframe
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        # Controle de janelas
        janelas_antes = navegador.window_handles

        # Localizar e clicar na lupa de marcas
        lupa_marca = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="marcas_btpesquisar"]/td[1]')))
        navegador.execute_script("arguments[0].scrollIntoView(true);", lupa_marca)
        lupa_marca.click()

        # Controle de janelas
        WebDriverWait(navegador, 10).until(
            lambda d: len(d.window_handles) > len(janelas_antes)
        )

        # Alternar para a nova janela
        nova_janela = [j for j in navegador.window_handles if j not in janelas_antes][0]
        navegador.switch_to.window(nova_janela)
        navegador.maximize_window()

        # Realizar a pesquisa e selecionar a marca
        espera.until(EC.element_to_be_clickable((By.ID, 'pesquisa'))).send_keys(marca)
        navegador.find_element(By.ID, 'sub').click()
        espera.until(EC.element_to_be_clickable((By.ID, f'chk_{id_marca}'))).click()
        navegador.find_element(By.ID, 'btFPAdicionar').click()

        # Voltar para a janela original
        navegador.switch_to.window(janelas_antes[0])
        print(f"Marca '{marca}' selecionada com sucesso!")

    def finalizar_relatorio(navegador, espera):

        """
        Finaliza a configuração do relatório e exporta para Excel.
        """
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        # Seleciona tipo de custo
        tipo_custo = Select(navegador.find_element(By.ID, 'tipo_custo'))
        tipo_custo.select_by_index(5)

        # Ajustes no relatório
        navegador.find_element(By.XPATH,
                               '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[28]/td[2]/input[2]').click()
        navegador.find_element(By.XPATH,
                               '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[30]/td[2]/input[3]').click()
        navegador.find_element(By.XPATH, '//*[@id="mostralistar"]/td[2]/input[2]').click()
        navegador.find_element(By.ID, 'exibir_barras').click()

        # Gerar o relatório
        navegador.find_element(By.CLASS_NAME, 'btn-submit').click()
        espera.until(EC.element_to_be_clickable((By.ID, 'botaoExportarXLS'))).click()

    def acessar_relatorio_novamente(navegador, espera, data_inicial, data_final):
        """
        Sai do iframe, volta ao menu de relatórios e acessa o relatório desejado.
        """
        # Sair do iframe e voltar ao conteúdo principal
        navegador.switch_to.default_content()
        print("Voltando ao menu principal fora do iframe...")

        # Clica no relatório desejado
        relatorio = espera.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="liModulo_10"]/ul/li[25]/ul/li[18]/a')))
        relatorio.click()

        # Voltar para o iframe após acessar o relatório
        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)
        print("Entrando no iframe novamente...")

        for campo, valor in [("f_data1", data_inicial), ("f_data2", data_final)]:
            campo_data = espera.until(EC.element_to_be_clickable((By.ID, campo)))
            campo_data.click()
            campo_data.send_keys(Keys.CONTROL + "a")
            campo_data.send_keys(Keys.BACKSPACE)
            campo_data.send_keys(valor)

        # Seleciona e desseleciona empresas
        navegador.find_element(By.ID, 'empresas_3').click()
        navegador.find_element(By.ID, 'empresas_2').click()

    def relatorio_lvmh(navegador, espera, data_inicial, data_final, fornecedor, cod_fornecedor, marca, cod_marca):

        # Voltar ao menu de relatórios e acessar novamente
        acessar_relatorio_novamente(navegador, espera, data_inicial, data_final)

        # Configurar o relatório para o fornecedor atual
        selecionar_fornecedor(navegador, espera, fornecedor, cod_fornecedor)

        time.sleep(2)

        # Seleciona a marca
        selecionar_marca(navegador, espera, marca, cod_marca)

        # Finalizar e salvar o relatório
        finalizar_relatorio(navegador, espera)
        time.sleep(3)  # Evitar sobreposição entre relatórios

    def relatorio_weitnauer(navegador, espera, data_inicial, data_final, fornecedor, cod_fornecedor):

        # Voltar ao menu de relatórios e acessar novamente
        acessar_relatorio_novamente(navegador, espera, data_inicial, data_final)

        # Configurar o relatório para o fornecedor atual
        selecionar_fornecedor(navegador, espera, fornecedor, cod_fornecedor)
        time.sleep(2)

        iframe = espera.until(EC.presence_of_element_located((By.ID, "main")))
        navegador.switch_to.frame(iframe)

        # Seleciona tipo de custo
        tipo_custo = Select(navegador.find_element(By.ID, 'tipo_custo'))
        tipo_custo.select_by_index(5)

        # Ajustes no relatório
        navegador.find_element(By.XPATH,
                               '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[28]/td[2]/input[2]').click()
        navegador.find_element(By.XPATH,
                               '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[30]/td[2]/input[8]').click()
        navegador.find_element(By.XPATH, '//*[@id="mostralistar"]/td[2]/input[2]').click()
        navegador.find_element(By.ID, 'exibir_barras').click()

        # Gerar o relatório
        navegador.find_element(By.CLASS_NAME, 'btn-submit').click()
        espera.until(EC.element_to_be_clickable((By.ID, 'botaoExportarXLS'))).click()

        time.sleep(3)

    def renomear_relatorios(pasta):
        """
        Renomeia os relatórios baixados e convertidos conforme o mapeamento especificado.
        """
        # Definir o mapeamento de nomes corrigindo possíveis espaços extras
        mapeamento = {
            "ProdutosServicosVendidos.xls": "PUIG-MARY.xls",
            "ProdutosServicosVendidos (1).xls": "DBB-MARY.xls",
            "ProdutosServicosVendidos (2).xls": "LOREAL-MARY.xls",
            "ProdutosServicosVendidos (3).xls": "EXCELLENCE-MARY.xls",
            "ProdutosServicosVendidos (4).xls": "PRESTIGE-MARY.xls",
            "ProdutosServicosVendidos (5).xls": "TABATINGA-MARY.xls",
            "ProdutosServicosVendidos (6).xls": "JUNO-MARY.xls",
            "ProdutosServicosVendidos (7).xls": "COTY-MARY.xls",
            "ProdutosServicosVendidos (8).xls": "WEITNAUER-MARY.xls",
            "ProdutosServicosVendidos (9).xls": "LVMH-MARY.xls"
        }

        # Confirma se a pasta existe
        if not os.path.exists(pasta):
            print(f"Erro: A pasta '{pasta}' não existe!")
            return

        # Lista os arquivos exatos na pasta
        arquivos = os.listdir(pasta)
        print(f"\n📂 Arquivos encontrados na pasta '{pasta}':")
        for arquivo in arquivos:
            print(f" - {arquivo}")

        # Renomear os arquivos
        for arquivo in arquivos:
            caminho_antigo = os.path.join(pasta, arquivo)

            # Verifica se o nome do arquivo está no mapeamento
            if arquivo in mapeamento:
                novo_nome = mapeamento[arquivo]
                caminho_novo = os.path.join(pasta, novo_nome)

                os.rename(caminho_antigo, caminho_novo)
                print(f"✅ Renomeado: {arquivo} → {novo_nome}")
            else:
                print(f"⚠ Arquivo ignorado: {arquivo} (não mapeado)")

    # Definir a pasta correta onde os arquivos foram baixados
    pasta_relatorios = "/Users/italosousa/Documents/Relatórios Fornecedores/Relatórios MARY"

    # ---------------- Fluxo Principal ----------------
    if __name__ == "__main__":

        # Lista de fornecedores (nome e ID)
        fornecedores = [
            {"nome": "DBB", "id": 127},
            {"nome": "loreal", "id": 224},
            {"nome": "EXCELLENCE", "id": 156},
            {"nome": "PRESTIGE", "id": 313},
            {"nome": "TABATINGA", "id": 373},
            {"nome": "JUNO", "id": 213},
            {"nome": "COTY", "id": 101494}
        ]

        navegador, espera = criar_navegador()

        try:
            # Credenciais
            fazer_login(navegador, espera, "italo.meca", "@86Kizzmacca6")
            selecionar_empresa(navegador, espera, 2)
            acessar_relatorio(navegador, espera)
            configurar_relatorio(navegador, espera, data_inicio, data_fim)
            selecionar_fornecedor(navegador, espera, "puig", 318)
            finalizar_relatorio(navegador, espera)
            time.sleep(3)

            for fornecedor in fornecedores:
                print(f"Gerando relatório para fornecedor: {fornecedor['nome']} (ID: {fornecedor['id']})")

                # Voltar ao menu de relatórios e acessar novamente
                acessar_relatorio_novamente(navegador, espera, data_inicio, data_fim)

                # Configurar o relatório para o fornecedor atual
                selecionar_fornecedor(navegador, espera, fornecedor["nome"], fornecedor["id"])

                # Finalizar e salvar o relatório
                finalizar_relatorio(navegador, espera)
                time.sleep(3)  # Evitar sobreposição entre relatórios

                print(f"Relatório para {fornecedor['nome']} gerado com sucesso!")

            # Relatório weitnauer que é agrupado por vendedor
            relatorio_weitnauer(navegador, espera, data_inicio, data_fim, "WEITNAUER", 405)

            # Gerando relatório LVMH que tem que selecionar marca
            relatorio_lvmh(navegador, espera, data_inicio, data_fim, "lvmh", 248, "dior", 6)

            # **Renomear os arquivos após a conversão**
            renomear_relatorios(pasta_relatorios)

            print("Todos os relatórios foram gerados, convertidos, renomeados e tratados com sucesso.")
        finally:
            navegador.quit()


# Criando a interface principal
root = tk.Tk()
root.title('Automações B.stories')
root.geometry("850x350")  # Ajustando tamanho para caber o terminal

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
geracaoRelatorios = tk.Button(left_frame, text='Geração Relatórios', fg='blue')
geracaoRelatorios.pack(pady=10, padx=15, fill="x")

ajusteEstoque = tk.Button(left_frame, text='Ajuste Estoque')
ajusteEstoque.pack(pady=10, padx=15, fill="x")

subirEstoque = tk.Button(left_frame, text='Subir Estoque')
subirEstoque.pack(pady=10, padx=15, fill="x")

# Texto sobre relatórios no content_area
datas = tk.Label(content_area, text='Insira as datas dos relatórios:', font='Helvetica 12 bold')
datas.grid(row=0, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

# Data inicial
dataInicial = tk.Label(content_area, text='Data inicial: ')
dataInicial.grid(row=1, column=1, padx=10, pady=5)

entryDataInicial = tk.Entry(content_area)
entryDataInicial.grid(row=2, column=1, padx=10, pady=5)

# Data final
dataFinal = tk.Label(content_area, text='Data final: ')
dataFinal.grid(row=3, column=1, padx=10, pady=5)

entryDataFinal = tk.Entry(content_area)
entryDataFinal.grid(row=4, column=1, padx=10, pady=5)

# Texto sobre relatórios no content_area
escolherBotoes = tk.Label(content_area, text='Selecione o relatório desejado:', font='Helvetica 12 bold')
escolherBotoes.grid(row=5, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

# Botão Gerar Relatório LD
botao_LD = tk.Button(content_area, text="Gerar Relatório LD",
                      command=lambda: validar_entrada(relatoriosLD, entryDataInicial.get(), entryDataFinal.get()))
botao_LD.grid(row=6, column=1, padx=10, pady=5)

# Botão Gerar Relatório MARY
botao_MARY = tk.Button(content_area, text="Gerar Relatório MARY",
                        command=lambda: validar_entrada(relatorioMARY, entryDataInicial.get(), entryDataFinal.get()))
botao_MARY.grid(row=7, column=1, padx=10, pady=5)


# Criando a área do "terminal" na interface
terminal_label = tk.Label(terminal_frame, text="Terminal", font="Helvetica 12 bold", fg="white")
terminal_label.pack(anchor="w", padx=10, pady=5)

# Criando um Frame extra branco para a borda do terminal
terminal_border = tk.Frame(terminal_frame, bg="white")
terminal_border.pack(padx=5, pady=5)

terminal_text = tk.Text(terminal_border, bg="black", fg="white", insertbackground="white", wrap="word")
terminal_text.pack(expand=True, fill="both", padx=10, pady=5)

# Redirecionando stdout para o widget Text
sys.stdout = RedirectText(terminal_text)

root.mainloop()