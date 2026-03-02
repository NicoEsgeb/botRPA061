from robocorp.tasks import task
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os 
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
from functions import read_excel_caratulas

EXCEL_INGRESO_DEMANDAS = "C:\\Applications\\RPA 06 - INGRESO DE DEMANDAS Y DOCUMENTOS EN PODER JUDICIAL\\input\\Itau_ddas_pjud\\BOT_MATRIZ_DEMANDAS.xlsx"
EXCEL_INFORME_PJUD = "C:\\Applications\\RPA 06 - INGRESO DE DEMANDAS Y DOCUMENTOS EN PODER JUDICIAL\\input\\Itau_ddas_pjud\\Informe_Causas_02.03.2026_13_27_40.xlsx"
    #"Informe_Causas_24.12.2025_16_09_19.xlsx"
CARATULAS_FOLDER_PATH = "C:\\Applications\\RPA 06 - INGRESO DE DEMANDAS Y DOCUMENTOS EN PODER JUDICIAL\\input\\Itau_ddas_pjud\\Caratulas"
FECHA_FILTRO_CARATULAS = "24/02/2026"

PATH_BOT = os.path.dirname(os.path.realpath(__file__))

def run_get_caratulas(
    excel_ingreso_demandas: str,
    excel_informe_pjud: str,
    caratulas_folder_path: str,
    fecha_filtro: str,
) -> None:
    """Ejecuta la descarga de caratulas usando parametros de entrada."""
    try:
        options = Options()
        options.add_argument('--start-maximized')

        default_directory = caratulas_folder_path
        
        options.add_experimental_option('prefs', {
            "download.default_directory": default_directory, #Change default directory for downloads
            "download.prompt_for_download": False, #To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
        })

        # Configurar el navegador
        chrome_install = ChromeDriverManager().install()

        folder = os.path.dirname(chrome_install)
        chromedriver_path = os.path.join(folder, "chromedriver.exe")

        cService = webdriver.ChromeService(chromedriver_path)
            
        # Crear el controlador
        driver = webdriver.Chrome(options=options, service = cService)

        driver.implicitly_wait(30)

        driver.get("https://ojv.pjud.cl/kpitec-ojv-web/views/login.html")

        try:
            driver.find_element(By.XPATH, "//*[@id='modalAviso']/div/div/div[1]/div/img").click()
        except:
            pass
        time.sleep(20)
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, 'link2C')))
        time.sleep(20)
        driver.find_element(By.ID, value='link2C').click()
        driver.find_element(By.ID, value='inputRut2C').send_keys('11346197')
        driver.find_element(By.ID, value='inputPassword2C').send_keys('Quinchilca33+')
        driver.find_element(By.CLASS_NAME, value='btn-ingreso-pjud').click()

        time.sleep(20)
        driver.find_element(By.CLASS_NAME, value='card-text').click()
        time.sleep(20)

        read_excel_caratulas(
            excel_informe_pjud,
            caratulas_folder_path,
            driver,
            excel_ingreso_demandas,
            fecha_filtro,
        )


    finally:
        # Place for teardown and cleanups
        # Playwright handles browser closing
        print('Done')

@task
def RPA_06_GET_CARATULAS() :
    """
    RPA 06 - Descargar las caratulas de las demandas en poder judicial.
    """
    run_get_caratulas(
        EXCEL_INGRESO_DEMANDAS,
        EXCEL_INFORME_PJUD,
        CARATULAS_FOLDER_PATH,
        FECHA_FILTRO_CARATULAS,
    )


@task
def RPA_06_DescargarInformePjud():
    """
    RPA 06 - Descarga del Excel del poder judicial
    
    """
    try:
        options = Options()
        options.add_argument('--start-maximized')

        default_directory = CARATULAS_FOLDER_PATH
        
        options.add_experimental_option('prefs', {
            "download.default_directory": default_directory, #Change default directory for downloads
            "download.prompt_for_download": False, #To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
        })

        # Configurar el navegador
        chrome_install = ChromeDriverManager().install()

        folder = os.path.dirname(chrome_install)
        chromedriver_path = os.path.join(folder, "chromedriver.exe")

        cService = webdriver.ChromeService(chromedriver_path)
            
        # Crear el controlador
        driver = webdriver.Chrome(options=options, service = cService)

        driver.implicitly_wait(30)
    
        driver.get("https://ojv.pjud.cl/kpitec-ojv-web/views/login.html")

        driver.find_element(By.ID, value='link2C').click()
        driver.find_element(By.ID, value='inputRut2C').send_keys('11346197')
        driver.find_element(By.ID, value='inputPassword2C').send_keys('Quinchilca33+')
        driver.find_element(By.CLASS_NAME, value='btn-ingreso-pjud').click()

        time.sleep(20)
        driver.find_element(By.CLASS_NAME, value='card-text').click()
        time.sleep(20)
        
        driver.get("https://ojv.pjud.cl/kpitec-ojv-web/index#bandeja/causas")        
        time.sleep(5)

        actions = ActionChains(driver)  

        driver.execute_script("document.body.style.zoom='70%'")       
                        
        driver.find_element(By.XPATH, "//label[contains(.,'Competencia:')]/../select").send_keys('Civil')
        actions.send_keys(Keys.TAB).perform()
        time.sleep(4)
        
        driver.find_element(By.XPATH, "//a[contains(.,'Demandas Enviadas')]").click()
        time.sleep(3)

        driver.find_element(By.XPATH, "//input[@value='Exportar a Excel']").click()
        time.sleep(15)

    finally:
        # Place for teardown and cleanups
        # Playwright handles browser closing
        #browser.close();
        #driver.quit()
        print('Done Descarga Informe Pjud')
