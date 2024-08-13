try:
    import os
    import pickle
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
except Exception as e:
    print(f"Erro ao importar Bibliotecas. Erro: {e}")

def configure_browser(download_dir, headless=False):
    current_dir = os.path.abspath(os.path.dirname(__file__))
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1270,720")
    if headless:
        chrome_options.add_argument("--headless")
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

def save_cookies(driver, cookies_file_path):
    with open(cookies_file_path, 'wb') as filehandler:
        pickle.dump(driver.get_cookies(), filehandler)

def login(driver, email, password):
    url = 'https://pt.semrush.com/login/?src=header&redirect_to=%2F'
    driver.get(url)

    try:
        permitir_cookies_button = WebDriverWait(driver, 100).until(
            EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Permitir Todos os Cookies")]'))
        )
        permitir_cookies_button.click()
        print("Botão 'Permitir Todos os Cookies' clicado.")
        
        email_field = WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.NAME, 'email'))
        )
        email_field.send_keys(email)
        
        password_field = WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.NAME, 'password'))
        )
        password_field.send_keys(password)
        
        login_button = WebDriverWait(driver, 100).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@type="submit"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView();", login_button)
        login_button.click()

        print("Aguardando CAPTCHA...")
        
        elemento_inicial = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, '//a[@href="https://pt.semrush.com/analytics/overview/?db=us" or @href="https://pt.semrush.com/analytics/overview/" or @href="https://pt.semrush.com/analytics/overview/?db=br"]'))
        )
        print(f"Elemento após login encontrado: {elemento_inicial}")
        print("Login realizado com sucesso e CAPTCHA resolvido.")

        return driver.get_cookies()
    except Exception as e:
        print(f"Erro ao tentar realizar o login: {e}")
        return None

def main():
    current_dir = os.path.abspath(os.path.dirname(__file__))
    download_dir = os.path.join(current_dir, "DWNLD")
    driver = configure_browser(download_dir, headless=False)
    
    email = 'your_email'
    password = 'your_password'

    cookies = login(driver, email, password)
    
    if cookies:
        save_cookies(driver, os.path.join(current_dir, 'cookies.pkl'))
        print("Cookies salvos com sucesso.")
    else:
        print("Login falhou. Não foi possível obter os cookies.")
    driver.quit()

if __name__ == "__main__":
    main()
