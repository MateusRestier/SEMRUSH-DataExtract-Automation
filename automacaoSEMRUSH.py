# UTILIZANDO VERSAO CHROME 127
# pip install selenium pandas openpyxl


try: # Tenta importar as bibliotecas necessárias
    import os  # Biblioteca para manipulação do sistema operacional
    import time  # Biblioteca para manipulação de tempo e pausas
    import pickle # Biblioteca para salvar e carregar cookies
    from selenium import webdriver  # Biblioteca Selenium para controle do navegador
    from selenium.webdriver.common.by import By  # Módulo para localizar elementos no navegador
    from selenium.webdriver.chrome.service import Service  # Módulo para iniciar o serviço do Chrome
    from selenium.webdriver.chrome.options import Options  # Módulo para configurar opções do Chrome
    from selenium.webdriver.common.keys import Keys  # Módulo para emular teclas do teclado
    from selenium.webdriver.support.ui import WebDriverWait  # Módulo para esperar elementos no navegador
    from selenium.webdriver.support import expected_conditions as EC  # Módulo para definir condições esperadas
except Exception as e:
    # Imprime mensagem de erro caso ocorra falha ao importar alguma biblioteca
    print(f"Erro ao importar Bibliotecas. Erro: {e}")


def configure_browser(download_dir, headless=False):
    current_dir = os.path.abspath(os.path.dirname(__file__)) # Caminho absoluto para o diretório atual
    if not os.path.exists(download_dir): # Garantir que o diretório de download exista, caso contrário, cria-o
        os.makedirs(download_dir) # Cria o diretório de download

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1366,720")
    chrome_options.add_argument("--log-level=3")  # Minimiza os logs, 3 corresponde a 'FATAL' (apenas erros críticos), assim o modo headlessn spamma informação
    chrome_options.add_argument("--headless") # Modo headless
        
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    
    chrome_driver_path = os.path.join(current_dir, 'chromedriver.exe')
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def load_cookies(driver, cookies_file_path): # Carrega os cookies de um arquivo e adiciona-os ao navegador
    driver.get('https://pt.semrush.com') # Acessa a página principal do site
    with open(cookies_file_path, 'rb') as cookiesfile: # Abre o arquivo de cookies
        cookies = pickle.load(cookiesfile) # Carrega os cookies do arquivo
        for cookie in cookies: # Adiciona cada cookie ao navegador
            driver.add_cookie(cookie) # Adiciona o cookie ao navegador
    driver.refresh()# Atualiza a página para aplicar os cookies


def wait_for_download_complete(download_dir, new_filename='download.csv', timeout=60):
    end_time = time.time() + timeout
    while time.time() < end_time:
        if any(fname.endswith('.crdownload') for fname in os.listdir(download_dir)):
            print("Arquivo .crdownload ainda presente. Esperando o download completar...")
            time.sleep(4)
        else:
            print("Arquivo .crdownload não encontrado. Verificando arquivos CSV...")
            if any(fname.endswith('.csv') for fname in os.listdir(download_dir)):
                print("Arquivo CSV encontrado.")
                if new_filename:
                    for fname in os.listdir(download_dir):
                        if fname.endswith('.csv'):
                            old_path = os.path.join(download_dir, fname)
                            new_path = os.path.join(download_dir, new_filename)
                            os.rename(old_path, new_path)
                            print(f"Renomeado {old_path} para {new_path}")
                return
            else:
                print("Arquivo CSV não encontrado. Continuando a espera...")
                time.sleep(1)
    raise Exception("Download não foi concluído dentro do tempo limite.")


def navegar_para_projetos(driver):
    """Navega para a aba 'Projetos'."""
    try:
        print("Esperando o link 'Projetos' aparecer...")
        projetos_link = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@href="/projects/"]'))
        )
        if projetos_link:
            print("Link 'Projetos' encontrado.")
            projetos_link.click()
        else:
            print("Link 'Projetos' não encontrado.")
            return
        print("Navegando para 'Projetos'.")
    except Exception as e:
        print(f"Erro ao navegar até a aba 'Projetos'. Erro: {e}")


def navegar_para_traffic_analytics(driver):
    """Navega para a aba 'Traffic Analytics'."""
    try:
        # Obtém o diretório atual do script e define o diretório de download
        current_dir = os.path.abspath(os.path.dirname(__file__))
        download_dir = os.path.join(current_dir, "DWNLD", "LPC")
        
        # Configura o comportamento de download para permitir downloads automáticos no diretório especificado
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': download_dir
        })
        print("Traffic Analytics' aparecer...")
        
        # Espera até que o link 'Traffic Analytics' esteja presente na página
        trafficanalytics_link = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//a[@href="https://pt.semrush.com/analytics/traffic/overview/" or @href="https://pt.semrush.com/analytics/traffic/overview/?db=us" or @href="https://pt.semrush.com/analytics/traffic/overview/?db=br"]'))
        )
        
        # Verifica se o link foi encontrado e clica nele
        if trafficanalytics_link:
            print("Link 'Traffic Analytics' encontrado.")
            trafficanalytics_link.click()
        else:
            print("Link 'Traffic Analytics' não encontrado.")
            return
        print("Traffic Analytics'.")
        
    except Exception as e:
        # Imprime um erro caso haja problema ao navegar para a aba 'Traffic Analytics'
        print(f"Erro ao navegar até a aba 'Traffic Analytics'. Erro: {e}")

def baixar_visao_geral_dominio(driver, domainbg):
    """Navega para a aba 'Visão geral do domínio' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    
    try:
        # Obtém o diretório atual do script e define o diretório de download
        current_dir = os.path.abspath(os.path.dirname(__file__))
        download_dir = os.path.join(current_dir, "DWNLD", "VGD")

        # Configura o comportamento de download para permitir downloads automáticos no diretório especificado
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': download_dir
        })

        print("Esperando o link 'Visão geral do domínio' aparecer...")
        visao_geral_dominio_link = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//a[@href="https://pt.semrush.com/analytics/overview/?db=us" or @href="https://pt.semrush.com/analytics/overview/" or @href="https://pt.semrush.com/analytics/overview/?db=br"]'))
        )

        if visao_geral_dominio_link:
            print("Link 'Visão geral do domínio' encontrado.")
            visao_geral_dominio_link.click()
        else:
            print("Link 'Visão geral do domínio' não encontrado.")
            return

        print("Navegando para 'Visão geral do domínio'.")

        print("Esperando a caixa de entrada do domínio aparecer...")
        domain_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Informe o domínio, subdomínio ou URL"]'))
        )
        print("Caixa de entrada do domínio encontrada.")
        
        domain_input.send_keys(domainbg)
        print("Domínio inserido na caixa de entrada.")
        
        print("Esperando o botão de pesquisa aparecer...")
        search_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="Analise o domínio"]'))
        )
        print("Botão de pesquisa encontrado.")
        
        search_button.click()
        print("Botão de pesquisa clicado.")
        
        # Esperar e clicar na opção "Meses"
        print("Esperando a opção 'Meses' aparecer...")
        meses_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-at="pill-monthly"]'))
        )
        driver.execute_script("arguments[0].click();", meses_option)
        print("Opção 'Meses' clicada.")
        
        print("Esperando o botão 'Exportar' aparecer...")
        export_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Exportar o relatório Tráfego"]'))
        )
        print("Botão 'Exportar' encontrado.")
        export_button.click()
        
        print("Esperando a opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@role="menuitem" and contains(text(),"CSV")]'))
        )
        print("Opção 'CSV' encontrada.")
        driver.execute_script("arguments[0].click();", csv_option)
        
        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão de pesquisa, botão 'Exportar' ou a opção 'CSV'. Erro: {e}")


def baixar_LacunasBacklinks(driver, domainbg, domainlp, domainin, domainses, domaingo):
    """Navega para a aba 'Lacunas nos Backlinks' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        # Obtém o diretório atual do script e define o diretório de download
        current_dir = os.path.abspath(os.path.dirname(__file__))
        download_dir = os.path.join(current_dir, "DWNLD", "LB")
        
        # Configura o comportamento de download para permitir downloads automáticos no diretório especificado
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': download_dir
        })
        
        print("Lacunas nos Backlinks' aparecer...")
        lacunasbacklink_link = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//a[@href="https://pt.semrush.com/analytics/gap/backlinks/" or @href="https://pt.semrush.com/analytics/gap/backlinks/?db=us" or @href="https://pt.semrush.com/analytics/gap/backlinks/?db=br"]'))
        ) 
        
        if lacunasbacklink_link:
            print("Link 'Lacunas nos Backlinks' encontrado.")
            lacunasbacklink_link.click()
        else:
            print("Link 'Lacunas nos Backlinks' não encontrado.")
            return
        print("Lacunas nos Backlinks'.")
        
        print("Esperando botão 'Acrescente até 3 concorrentes' aparecer...")
        adicionar_concorrentes_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-test-add-competitors]'))
        )
        adicionar_concorrentes_button.click()
        print("Botão 'Acrescente até 3 concorrentes' clicado.")
        
        print("Esperando caixas de entrada para domínios aparecerem...")
        caixas_de_entrada = WebDriverWait(driver, 60).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'input[placeholder="Domínio"]'))
        )
        
        # Verifica se há pelo menos 5 caixas de entrada e insere os domínios
        if len(caixas_de_entrada) >= 5:
            caixas_de_entrada[0].send_keys(domainbg)
            caixas_de_entrada[1].send_keys(domainlp)
            caixas_de_entrada[2].send_keys(domainin)
            caixas_de_entrada[3].send_keys(domainses)
            caixas_de_entrada[4].send_keys(domaingo)
            print("Domínios inseridos nas caixas de entrada.")
        else:
            print("Não foi possível encontrar todas as caixas de entrada.")
            return
        
        print("Esperando botão 'Identificar perspectivas' aparecer...")
        identificar_perspectivas_button = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//button[@data-test-target-selector-submit and @data-ui-name="Button"]'))
        )
        driver.execute_script("arguments[0].click();", identificar_perspectivas_button)
        print("Botão 'Identificar perspectivas' clicado.")
        
        # Clicar no botão "Todos" antes de exportar
        print("Esperando botão 'Todos' aparecer...")
        todos_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-test-filter-preset="all" and @value="all"]'))
        )
        todos_button.click()
        print("Botão 'Todos' clicado.")
        
        print("Esperando botão 'Exportação' aparecer...")
        exportacao_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@id="igc-ui-kit-r2c-trigger"]'))
        )
        exportacao_button.click()
        print("Botão 'Exportação' clicado.")
        
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@data-test-export-type="csv"]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")

        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão de pesquisa, botão 'Exportar' ou a opção 'CSV'. Erro: {e}")



def baixar_LacunasPalavrasChave(driver, domainbg, domainlp, domainin, domainses, domaingo): 
    """Navega para a aba 'Lacunas nas palavras-chave' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        current_dir = os.path.abspath(os.path.dirname(__file__))
        download_dir = os.path.join(current_dir, "DWNLD", "LPC")
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': download_dir
        })
        print("Esperando 'Lacunas nas palavras-chave' aparecer...")
        lacunaspalavrachave_link = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//a[@href="https://pt.semrush.com/analytics/keywordgap/?db=us" or @href="https://pt.semrush.com/analytics/keywordgap/" or @href="https://pt.semrush.com/analytics/keywordgap/?db=br"]'))
        )
        if lacunaspalavrachave_link:
            print("Link 'Lacunas nas palavras-chave' encontrado.")
            lacunaspalavrachave_link.click()
        else:
            print("Link 'Lacunas nas palavras-chave' não encontrado.")
            return
        print("Lacunas nas palavras-chave'.")
        
        print("Esperando botão 'Acrescente até 3 concorrentes' aparecer...")
        adicionar_concorrentes_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-at="add-competitors-btn add-3"]'))
        )
        adicionar_concorrentes_button.click()
        print("Botão 'Acrescente até 3 concorrentes' clicado.")
        
        print("Esperando caixas de entrada para domínios aparecerem...")
        caixas_de_entrada = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'input[data-at="competitor-value"]'))
        )
        
        # Verifica se há pelo menos 5 caixas de entrada e insere os domínios
        if len(caixas_de_entrada) >= 5:
            caixas_de_entrada[0].send_keys(domainbg)
            caixas_de_entrada[1].send_keys(domainlp)
            caixas_de_entrada[2].send_keys(domainin)
            caixas_de_entrada[3].send_keys(domainses)
            caixas_de_entrada[4].send_keys(domaingo)
            print("Domínios inseridos nas caixas de entrada.")
        else:
            print("Não foi possível encontrar todas as caixas de entrada.")
            return
        
        print("Esperando botão 'Comparar' aparecer...")
        comparar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-at="compare-btn"]'))
        )
        driver.execute_script("arguments[0].click();", comparar_button)
        print("Botão 'Comparar' clicado.")

        # Esperar e rolar a página até o botão "Exportar"
        print("Esperando botão 'Exportar' aparecer...")
        exportar_button = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//button[contains(@aria-label, "Exportar os detalhes das palavras-chave")]'))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", exportar_button)
        time.sleep(2)  # Espera 2 segundos para garantir que a rolagem seja concluída
        exportar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Exportar os detalhes das palavras-chave")]'))
        )
        driver.execute_script("arguments[0].click();", exportar_button)
        print("Botão 'Exportar' clicado.")
        
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Exportar para CSV")]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")

        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão de pesquisa, botão 'Exportar' ou a opção 'CSV'. Erro: {e}")

def baixar_TaVisitasSite(driver, domainbg, domainlp, domainin, domainses, domaingo):
    """Navega para a aba 'Visitas nos Sites (Visão Geral)' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        # Navegar para a aba 'Traffic Analytics'
        navegar_para_traffic_analytics(driver)
        
        current_dir = os.path.abspath(os.path.dirname(__file__))  # Diretório onde o script está localizado
        download_dir = os.path.join(current_dir, "DWNLD", "TaVS")  # Diretório de download
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {  # Configurar o comportamento de download
            'behavior': 'allow',  # Permitir downloads no diretório especificado
            'downloadPath': download_dir  # Diretório de download
        })
        
        # Esperar e clicar na caixa de entrada
        print("Esperando caixa de entrada 'Informe domínios, subdomínios ou subpastas' aparecer...")
        caixa_entrada = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[placeholder="Informe domínios, subdomínios ou subpastas"]'))
        )
        caixa_entrada.click()
        
        # Inserir domínios com intervalo de 1 segundo entre cada um
        dominios = [domainbg, domainlp, domainin, domainses, domaingo]
        for dominio in dominios:
            caixa_entrada.send_keys(dominio)
            caixa_entrada.send_keys(Keys.RETURN)
            time.sleep(1)
        
        print("Domínios inseridos na caixa de entrada.")
        
        # Esperar e clicar no botão "Analisar" com JavaScript
        print("Esperando botão 'Analisar' aparecer...")
        analisar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-test="submit" and @data-test-submit-type="analyze"]'))
        )
        driver.execute_script("arguments[0].click();", analisar_button)
        print("Botão 'Analisar' clicado.")
        
        time.sleep(2)  # Espera adicional para garantir que a página carregue completamente
        
        # Rolar para baixo para garantir que o botão "Exportar" esteja visível
        print("Rolando a página para encontrar o botão 'Exportar'...")
        driver.execute_script("window.scrollBy(0, 400);")
        
        # Esperar e clicar no botão "Exportar" específico
        print("Esperando botão 'Exportar' aparecer...")
        export_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="export-trigger" and contains(@class, "___SBoxInline_orfji_gg_") and .//span[text()="Exportar"]]'))
        )
        driver.execute_script("arguments[0].click();", export_button)
        print("Botão 'Exportar' clicado.")
        
        # Esperar a opção "CSV" aparecer e clicar nela
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-testid="export-csv"]/ancestor::div[@role="menuitem"]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")

        # Esperar o download ser concluído
        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão 'Analisar' ou botão 'Exportar'. Erro: {e}")


def baixar_TaTaxaRejeicao(driver, domainbg, domainlp, domainin, domainses, domaingo):
    """Navega para a aba 'Taxa de Rejeição' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        # Navegar para a aba 'Traffic Analytics'
        navegar_para_traffic_analytics(driver)
        
        current_dir = os.path.abspath(os.path.dirname(__file__))  # Diretório onde o script está localizado
        download_dir = os.path.join(current_dir, "DWNLD", "TaTR")  # Diretório de download
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {  # Configurar o comportamento de download
            'behavior': 'allow',  # Permitir downloads no diretório especificado
            'downloadPath': download_dir  # Diretório de download
        })
        
        # Esperar e clicar na caixa de entrada
        print("Esperando caixa de entrada 'Informe domínios, subdomínios ou subpastas' aparecer...")
        caixa_entrada = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[placeholder="Informe domínios, subdomínios ou subpastas"]'))
        )
        caixa_entrada.click()
        
        # Inserir domínios com intervalo de 1 segundo entre cada um
        dominios = [domainbg, domainlp, domainin, domainses, domaingo]
        for dominio in dominios:
            caixa_entrada.send_keys(dominio)
            caixa_entrada.send_keys(Keys.RETURN)
            time.sleep(1)
        
        print("Domínios inseridos na caixa de entrada.")
        
        # Esperar e clicar no botão "Analisar" com JavaScript
        print("Esperando botão 'Analisar' aparecer...")
        analisar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-test="submit" and @data-test-submit-type="analyze"]'))
        )
        driver.execute_script("arguments[0].click();", analisar_button)
        print("Botão 'Analisar' clicado.")
        
        # Esperar a página carregar completamente
        print("Esperando a página carregar completamente...")
        time.sleep(3)
        
        print("Rolando a página para encontrar o botão '...'.")
        driver.execute_script("window.scrollBy(0, 400);")
        
        # Esperar e clicar no botão de três pontinhos
        print("Esperando botão de três pontinhos aparecer...")
        tres_pontinhos_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@data-ui-name="Select.Trigger" and @aria-label="Outras métricas"]'))
        )
        driver.execute_script("arguments[0].click();", tres_pontinhos_button)
        print("Botão de três pontinhos clicado.")
        time.sleep(3)
        # Esperar e clicar na opção "Taxa de Rejeição"
        print("Esperando opção 'Taxa de Rejeição' aparecer...")
        taxa_rejeicao_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@data-ui-name="Select.Option" and @data-testid="bounceRate"]'))
        )
        driver.execute_script("arguments[0].click();", taxa_rejeicao_option)
        print("Opção 'Taxa de Rejeição' clicada.")
        
        # Esperar e clicar no botão "Exportar"
        print("Esperando botão 'Exportar' aparecer...")
        exportar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-ui-name="DropdownMenu.Trigger" and @data-testid="export-trigger"]'))
        )
        driver.execute_script("arguments[0].click();", exportar_button)
        print("Botão 'Exportar' clicado.")
        
        # Esperar e clicar na opção "CSV"
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-testid="export-csv"]/ancestor::div[@role="menuitem"]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")
        
        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão 'Analisar', botão de três pontinhos, opção 'Taxa de Rejeição' ou botão 'Exportar'. Erro: {e}")


def baixar_TaMediaDuracaoVisita(driver, domainbg, domainlp, domainin, domainses, domaingo):
    """Navega para a aba 'Media duração das visitas' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        # Navegar para a aba 'Traffic Analytics'
        navegar_para_traffic_analytics(driver)
        
        current_dir = os.path.abspath(os.path.dirname(__file__))  # Diretório onde o script está localizado
        download_dir = os.path.join(current_dir, "DWNLD", "TaMDV")  # Diretório de download
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {  # Configurar o comportamento de download
            'behavior': 'allow',  # Permitir downloads no diretório especificado
            'downloadPath': download_dir  # Diretório de download
        })
        
        # Esperar e clicar na caixa de entrada
        print("Esperando caixa de entrada 'Informe domínios, subdomínios ou subpastas' aparecer...")
        caixa_entrada = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[placeholder="Informe domínios, subdomínios ou subpastas"]'))
        )
        caixa_entrada.click()
        
        # Inserir domínios com intervalo de 1 segundo entre cada um
        dominios = [domainbg, domainlp, domainin, domainses, domaingo]
        for dominio in dominios:
            caixa_entrada.send_keys(dominio)
            caixa_entrada.send_keys(Keys.RETURN)
            time.sleep(1)
        
        print("Domínios inseridos na caixa de entrada.")
        
        # Esperar e clicar no botão "Analisar" com JavaScript
        print("Esperando botão 'Analisar' aparecer...")
        analisar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-test="submit" and @data-test-submit-type="analyze"]'))
        )
        driver.execute_script("arguments[0].click();", analisar_button)
        print("Botão 'Analisar' clicado.")
        
        # Esperar a página carregar completamente
        print("Esperando a página carregar completamente...")
        time.sleep(3)
        
        print("Rolando a página para encontrar o botão '...'.")
        driver.execute_script("window.scrollBy(0, 400);")
        
        # Esperar e clicar no botão de três pontinhos
        print("Esperando botão de três pontinhos aparecer...")
        tres_pontinhos_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@data-ui-name="Select.Trigger" and @aria-label="Outras métricas"]'))
        )
        driver.execute_script("arguments[0].click();", tres_pontinhos_button)
        print("Botão de três pontinhos clicado.")
        
        # Esperar e clicar na opção "Duração méd. da visita"
        print("Esperando opção 'Duração méd. da visita' aparecer...")
        duracao_media_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@data-ui-name="Select.Option" and @data-testid="timeOnSite"]'))
        )
        driver.execute_script("arguments[0].click();", duracao_media_option)
        print("Opção 'Duração méd. da visita' clicada.")
        
        # Esperar e clicar no botão "Exportar"
        print("Esperando botão 'Exportar' aparecer...")
        exportar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-ui-name="DropdownMenu.Trigger" and @data-testid="export-trigger"]'))
        )
        driver.execute_script("arguments[0].click();", exportar_button)
        print("Botão 'Exportar' clicado.")
        
        # Esperar e clicar na opção "CSV"
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-testid="export-csv"]/ancestor::div[@role="menuitem"]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")
        
        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão 'Analisar', botão de três pontinhos, opção 'Duração méd. da visita' ou botão 'Exportar'. Erro: {e}")

def baixar_TaJornadaTrafego(driver, domainbg, domainlp, domainin, domainses, domaingo):
    """Navega para a aba 'Jornada de Tréfego' e faz o download do CSV e PNG."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        # Navegar para a aba 'Traffic Analytics'
        navegar_para_traffic_analytics(driver)
        
        current_dir = os.path.abspath(os.path.dirname(__file__))  # Diretório onde o script está localizado
        download_dir = os.path.join(current_dir, "DWNLD", "TaJT")  # Diretório de download
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {  # Configurar o comportamento de download
            'behavior': 'allow',  # Permitir downloads no diretório especificado
            'downloadPath': download_dir  # Diretório de download
        })
        
        # Esperar e clicar na caixa de entrada
        print("Esperando caixa de entrada 'Informe domínios, subdomínios ou subpastas' aparecer...")
        caixa_entrada = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[placeholder="Informe domínios, subdomínios ou subpastas"]'))
        )
        caixa_entrada.click()
        
        # Inserir domínios com intervalo de 1 segundo entre cada um
        dominios = [domainbg, domainlp, domainin, domainses, domaingo]
        for dominio in dominios:
            caixa_entrada.send_keys(dominio)
            caixa_entrada.send_keys(Keys.RETURN)
            time.sleep(1)
        
        print("Domínios inseridos na caixa de entrada.")
        
        # Esperar e clicar no botão "Analisar" com JavaScript
        print("Esperando botão 'Analisar' aparecer...")
        analisar_button = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//button[@data-test="submit" and @data-test-submit-type="analyze"]'))
        )
        driver.execute_script("arguments[0].click();", analisar_button)
        print("Botão 'Analisar' clicado.")
        
        # Esperar e clicar na aba "Jornada de tráfego"
        print("Esperando aba 'Jornada de tráfego' aparecer...")
        jornada_trafego_tab = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@data-crop-value="Jornada de tráfego"]'))
        )
        driver.execute_script("arguments[0].click();", jornada_trafego_tab)
        print("Aba 'Jornada de tráfego' clicada.")
        
        # Esperar e clicar no botão "Exportar" com JavaScript
        print("Esperando botão 'Exportar' aparecer...")
        export_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-ui-name="DropdownMenu.Trigger" and @data-testid="export-trigger"]'))
        )
        driver.execute_script("arguments[0].click();", export_button)
        print("Botão 'Exportar' clicado.")
        
        # Esperar a opção "CSV" aparecer e clicar nela
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-testid="export-csv"]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")

        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download do CSV concluído.")

        time.sleep(3)
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão 'Analisar', aba 'Jornada de tráfego', botão 'Exportar' ou opção 'CSV'. Erro: {e}")

def baixar_VisaoGeralPalavrasChave(driver, personagens):
    """Navega para a aba 'Visao Geral Palavras Chave' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        current_dir = os.path.abspath(os.path.dirname(__file__)) # Diretório onde o script está localizado
        download_dir = os.path.join(current_dir, "DWNLD", "VGPC") # Diretório de download
        driver.execute_cdp_cmd('Page.setDownloadBehavior', { # Configurar o comportamento de download
            'behavior': 'allow', # Permitir downloads no diretório especificado
            'downloadPath': download_dir # Diretório de download
        })
        print("Visao Geral Palavras Chave' aparecer...")

        # Ajustar o seletor XPath para encontrar o link "Visão geral de palavras-chave"
        visaogeralpalavraschave_link = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@href="https://pt.semrush.com/analytics/keywordoverview/?db=us" or @href="https://pt.semrush.com/analytics/keywordoverview/" or @href="https://pt.semrush.com/analytics/keywordoverview/?db=br"]'))
        )
        if visaogeralpalavraschave_link:
            print("Link 'Visao Geral Palavras Chave' encontrado.")
            driver.execute_script("arguments[0].click();", visaogeralpalavraschave_link)
        else:
            print("Link 'Visao Geral Palavras Chave' não encontrado.")
            return
        print("Visao Geral Palavras Chave'.")

        # Esperar a página carregar e a caixa de entrada estar disponível
        print("Esperando caixa de entrada 'Informe a palavra-chave' aparecer...")
        caixa_entrada = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@role="textbox" and @class="kwo-start-page-textarea__textarea"]'))
        )

        # Garantir que a caixa de entrada esteja focada e visível
        driver.execute_script("arguments[0].scrollIntoView();", caixa_entrada)
        time.sleep(1)  # Aguardar um curto período para garantir que a ação seja completada
        driver.execute_script("arguments[0].click();", caixa_entrada)

        # Inserir o conteúdo da variável 'personagens' na caixa de entrada
        caixa_entrada.send_keys(personagens)
        print("Conteúdo da variável 'personagens' inserido na caixa de entrada.")

        # Esperar e clicar no botão "Pesquisar" com JavaScript
        print("Esperando botão 'Pesquisar' aparecer...")
        pesquisar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="search-button"]'))
        )
        driver.execute_script("arguments[0].click();", pesquisar_button)
        print("Botão 'Pesquisar' clicado.")

        # Esperar 10 segundos para garantir que a página carregue completamente
        print("Contagem Regressiva para esperar página carregar completamente")

        for i in range(15, -1, -1):
            print(i)
            time.sleep(1)               
        print("Página carregada.")

        # Esperar e clicar no botão "Exportar" com JavaScript
        print("Esperando botão 'Exportar' aparecer...")
        exportar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-ui-name="Dropdown.Trigger" and @data-testid="export-bulk-button"]'))
        )
        driver.execute_script("arguments[0].click();", exportar_button)
        print("Botão 'Exportar' clicado.")

        # Esperar e clicar na opção "CSV"
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-ui-name="Button.Text" and text()="CSV"]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")

        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão de pesquisa, botão 'Exportar' ou a opção 'CSV'. Erro: {e}")


def baixar_VisaoGeralPalavrasChave2(driver, personagens2):
    """Navega para a aba 'Visao Geral Palavras Chave' e faz o download do CSV."""
    
    # Navegar para a aba 'Projetos'
    navegar_para_projetos(driver)
    try:
        current_dir = os.path.abspath(os.path.dirname(__file__)) # Diretório onde o script está localizado
        download_dir = os.path.join(current_dir, "DWNLD", "VGPC2") # Diretório de download
        driver.execute_cdp_cmd('Page.setDownloadBehavior', { # Configurar o comportamento de download
            'behavior': 'allow', # Permitir downloads no diretório especificado
            'downloadPath': download_dir # Diretório de download
        })
        print("Visao Geral Palavras Chave' aparecer...")

        # Ajustar o seletor XPath para encontrar o link "Visão geral de palavras-chave"
        visaogeralpalavraschave2_link = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@href="https://pt.semrush.com/analytics/keywordoverview/?db=us" or @href="https://pt.semrush.com/analytics/keywordoverview/" or @href="https://pt.semrush.com/analytics/keywordoverview/?db=br"]'))
        )
        if visaogeralpalavraschave2_link:
            print("Link 'Visao Geral Palavras Chave' encontrado.")
            driver.execute_script("arguments[0].click();", visaogeralpalavraschave2_link)
        else:
            print("Link 'Visao Geral Palavras Chave' não encontrado.")
            return
        print("Visao Geral Palavras Chave'.")

        # Esperar a página carregar e a caixa de entrada estar disponível
        print("Esperando caixa de entrada 'Informe a palavra-chave' aparecer...")
        caixa_entrada = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@role="textbox" and @class="kwo-start-page-textarea__textarea"]'))
        )

        # Garantir que a caixa de entrada esteja focada e visível
        driver.execute_script("arguments[0].scrollIntoView();", caixa_entrada)
        time.sleep(1)  # Aguardar um curto período para garantir que a ação seja completada
        driver.execute_script("arguments[0].click();", caixa_entrada)

        # Inserir o conteúdo da variável 'personagens2' na caixa de entrada
        caixa_entrada.send_keys(personagens2)
        print("Conteúdo da variável 'personagens2' inserido na caixa de entrada.")

        # Esperar e clicar no botão "Pesquisar" com JavaScript
        print("Esperando botão 'Pesquisar' aparecer...")
        pesquisar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="search-button"]'))
        )
        driver.execute_script("arguments[0].click();", pesquisar_button)
        print("Botão 'Pesquisar' clicado.")

        # Esperar 10 segundos para garantir que a página carregue completamente
        print("Contagem Regressiva para esperar página carregar completamente")

        for i in range(15, -1, -1):
            print(i)
            time.sleep(1)               
        print("Página carregada.")

        # Esperar e clicar no botão "Exportar" com JavaScript
        print("Esperando botão 'Exportar' aparecer...")
        exportar_button = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@data-ui-name="Dropdown.Trigger" and @data-testid="export-bulk-button"]'))
        )
        driver.execute_script("arguments[0].click();", exportar_button)
        print("Botão 'Exportar' clicado.")

        # Esperar e clicar na opção "CSV"
        print("Esperando opção 'CSV' aparecer...")
        csv_option = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-ui-name="Button.Text" and text()="CSV"]'))
        )
        driver.execute_script("arguments[0].click();", csv_option)
        print("Opção 'CSV' selecionada.")

        print(f"Esperando o download ser concluído na pasta {download_dir}...")
        wait_for_download_complete(download_dir)
        print("Download concluído.")
        
    except Exception as e:
        print(f"Erro ao localizar ou interagir com a caixa de entrada, botão de pesquisa, botão 'Exportar' ou a opção 'CSV'. Erro: {e}")


def main():
    current_dir = os.path.abspath(os.path.dirname(__file__))
    download_dir = os.path.join(current_dir, "DWNLD")
    driver = configure_browser(download_dir, headless=False)
    
    # Carrega cookies de sessão salvos do primeiro script
    cookies_file_path = os.path.join(current_dir, 'cookies.pkl')
    load_cookies(driver, cookies_file_path)
    
    try:
        domainbg = 'www.bagaggio.com.br'
        domainlp = 'www.lepostiche.com.br'
        domainin = 'www.inovathi.com.br'
        domainses = 'www.sestini.com.br'
        domaingo = 'www.gocase.com.br'
        personagens = 'ARIEL,BRANCA DE NEVE,PATRULHA CANINA,CARROS,CAPITAO AMERICA,FROZEN,FUTEBOL,DINOSSAURO,HOMEM ARANHA,MARIE,MICKEY,MINNIE,MOANA,PRINCESAS,PETS,REI LEAO,SPIDEY,STITCH,TOY STORY,UNICORNIO,VINGADORES,FLAMENGO,POP FUN,ENALDINHO,REBECCA BONBON,CINDERELA,JASMINE,SONIC,PEPPA PIG,ROBLOX,MINECRAFT,POCOYO,GALINHA PINTADINHA,MINIONS,BABY SHARK,MORANGUINHO,BLUEY,SUPER MAN,BATMAN,ONE PIECE,LUCAS NETO,LULUCA,PANTERA NEGRA,PRINCESA SOFIA,BOLOFOFOS,HARRY POTTER,NARUTO,VASCO,BOTAFOGO,FLUMINENSE,CORINTHIANS,PALMEIRAS,SÃO PAULO,RB BRAGANTINO,ATLÉTICO,CRUZEIRO,INTERNACIONAL,GRÊMIO'
        personagens2 = 'SANTOS,MONICA,TURMA DA MONICA,WANDINHA,PJ MASK,PLAYSTATION,DORA AVENTUREIRA,BOB ESPONJA,GATO GALÁCTICO,MARIA CLARA E JP,FREE FIRE,POKEMON,STRANGER THINGS,SUPER MARIO,MARIO,URSINHOS CARINHOSOS,COCOMELON,AUTHENTIC GAMES,MUNDO BITA,PKXD,MY LITTLE PONY,BABY ALIVE,AMONG US,BARBIE,LOL,HOT WHEELS,FISHER PRICE,POLLY POCKET,MR POTATO HEAD,PLAY DOH,TARTARUGAS NINJAS,MTV,RICK AND MORTY,LIVERPOOL,ARSENAL,FLASH,MULHER MARAVILHA,FIFA,MENINAS SUPERPODEROSAS,LIGA DA JUSTIÇA,DC ORIGINALS,DC SUPER FRIENDS,DC SUPER HERO GIRLS,SHREK,A CASA MÁGICA DA GABY,JURASSIC WORLD,ONDE ESTÁ WALLY,TROLLS,BARCELONA,BAYERN DE MUNIQUE,PARIS SAINT GERMAIN,MANCHESTER CITY,NFL,PAC MAN,TURMA DA MATA,TURMA DO NEYMAR JR.,SIMPSONS,NEYMAR JR.,BRASIL'
        
        baixar_TaVisitasSite(driver, domainbg, domainlp, domainin, domainses, domaingo), time.sleep(3) 
        baixar_visao_geral_dominio(driver, domainbg), time.sleep(3)
        baixar_LacunasBacklinks(driver, domainbg, domainlp, domainin, domainses, domaingo), time.sleep(3)
        baixar_LacunasPalavrasChave(driver, domainbg, domainlp, domainin, domainses, domaingo), time.sleep(3)
        baixar_TaJornadaTrafego(driver, domainbg, domainlp, domainin, domainses, domaingo), time.sleep(3)
        baixar_VisaoGeralPalavrasChave(driver, personagens), time.sleep(3)
        baixar_VisaoGeralPalavrasChave2(driver, personagens2), time.sleep(3)
        baixar_TaMediaDuracaoVisita(driver, domainbg, domainlp, domainin, domainses, domaingo), time.sleep(3)
        baixar_TaTaxaRejeicao(driver, domainbg, domainlp, domainin, domainses, domaingo), time.sleep(3)
        
    finally:
        input("Pressione Enter para sair...")
        driver.quit()

if __name__ == "__main__":
    main()
