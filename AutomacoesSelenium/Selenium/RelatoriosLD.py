from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import pandas as pd
import os
import time

# Seleciona as data de inicio e de fim de acordo como ususario
data_inicio = input("Digite a data de início (formato DD/MM/AAAA): ")
data_fim = input("Digite a data de fim (formato DD/MM/AAAA): ")


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
    navegador.find_element(By.XPATH, '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[28]/td[2]/input[2]').click()
    navegador.find_element(By.XPATH, '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[30]/td[2]/input[3]').click()
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
    navegador.find_element(By.XPATH, '//*[@id="Form1"]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[28]/td[2]/input[2]').click()
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

