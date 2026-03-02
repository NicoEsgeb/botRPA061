from RPA.Excel.Files import Files as Excel
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from pypdf import PdfWriter
import glob

def read_excel_caratulas(
    EXCEL_INFORME_PJUD: str,
    CARATULAS_FOLDER_PATH: str,
    driver,
    EXCEL_INGRESO_DEMANDAS,
    fecha_filtro: str,
):
    """
    open and read an excel file for get caratulas

    """
    lista_ordenada = sorted(glob.glob(CARATULAS_FOLDER_PATH+"\*.pdf"))
    for file in lista_ordenada :
        print(file)

    print('Caratulas')
    excel = Excel()

    if os.path.isabs(EXCEL_INFORME_PJUD):
        path = EXCEL_INFORME_PJUD
    else:
        path = os.path.join(CARATULAS_FOLDER_PATH, EXCEL_INFORME_PJUD)

    excel.open_workbook(path)
    rows = excel.read_worksheet_as_table("Demandas Enviadas", header=True, start=4)

    caratulaIndex = 1;

    if not rows.columns.__contains__("DESCARGA_CARATULA") :
        excel.insert_columns_before(1,1)     
        excel.set_cell_value(4,1,"DESCARGA_CARATULA")
        excel.save_workbook()
    else :
        caratulaIndex = rows.columns.index("DESCARGA_CARATULA") + 1

    rows = excel.read_worksheet_as_table("Demandas Enviadas", header=True, start=4)

    tribunalesList = []
    fechaEnvio = ''
        
    for row in rows:
        if (not tribunalesList.__contains__(row["Tribunal"]) and row["DESCARGA_CARATULA"] != "OK") :
            tribunalesList.append(row["Tribunal"])
            fechaEnvio = str(row["Fecha Envio"])

    for tribunal in tribunalesList:

        driver.get("https://ojv.pjud.cl/kpitec-ojv-web/index#bandeja/causas")        
        time.sleep(5)

        driver.find_element(By.XPATH, "//label[contains(.,'Fecha Desde')]/../div/input").clear()
        # driver.find_element(By.XPATH, "//label[contains(.,'Fecha Desde')]/../div/input").send_keys(fechaEnvio)
        driver.find_element(By.XPATH, "//label[contains(.,'Fecha Desde')]/../div/input").send_keys(fecha_filtro)

        driver.find_element(By.XPATH, "//label[contains(.,'Fecha Hasta')]/../div/input").clear()
        # driver.find_element(By.XPATH, "//label[contains(.,'Fecha Hasta')]/../div/input").send_keys(fechaEnvio)
        driver.find_element(By.XPATH, "//label[contains(.,'Fecha Hasta')]/../div/input").send_keys(fecha_filtro)
        time.sleep(2)

        actions = ActionChains(driver)  

        driver.execute_script("document.body.style.zoom='70%'")
        
        index = 0
        tribunal_pantalla = ''
        while tribunal_pantalla != tribunal or index < 10 :                        
            driver.find_element(By.XPATH, "//label[contains(.,'Competencia:')]/../select").send_keys('Civil')
            time.sleep(4)
                  
            actions.send_keys(Keys.TAB).perform()            

            actions.send_keys(tribunal).perform()

            num_tribunal = tribunal[0:tribunal.index("º")]

            if int(num_tribunal) == 2:                
                actions.send_keys(Keys.ARROW_DOWN).perform()
            elif int(num_tribunal) > 2 and int(num_tribunal) < 10:
                actions.send_keys(Keys.ARROW_DOWN * 2).perform()

            time.sleep(2)
            actions.send_keys(Keys.TAB).perform()

            driver.find_element(By.XPATH, "//a[contains(.,'Demandas Enviadas')]").click()
            time.sleep(3)
            index += 1
            tribunal_pantalla = driver.find_element(By.XPATH, "//*[@id='idTrCardBandCau-0']/div/div/div[2]/div[2]/div[1]/span").text
            if tribunal_pantalla == tribunal :
                break

        demandas =  len(driver.find_elements(By.XPATH, '//*[@id="idBandejaCausa"]/div/div'))

        for x in range(demandas) :  

            if (x % 5 == 0 and x > 0) :
                actions.send_keys(Keys.ARROW_DOWN * 6).perform() 

            time.sleep(2)
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//*[@id='idTrCardBandCau-"+str(x)+"']/div/div/div[1]/a")))

            nombre_original = driver.find_element(By.XPATH, "//*[@id='idTrCardBandCau-"+str(x)+"']/div/div/div[2]/div[2]/div[5]/span").text.split("-")[1] + ".pdf"
            nombre_original = 'getDoc.pdf'
            
            ruta_original = os.path.join(CARATULAS_FOLDER_PATH, nombre_original)

            if os.path.exists(ruta_original):
                os.remove(ruta_original)
            
            #Descargo
            driver.find_element(By.XPATH, "//*[@id='idTrCardBandCau-"+str(x)+"']/div/div/div[1]/a").click()   
            time.sleep(1)

            rit = driver.find_element(By.XPATH, "//*[@id='idTrCardBandCau-"+str(x)+"']/div/div/div[2]/div[2]/div[4]/span").text

            ruta_caratula = CARATULAS_FOLDER_PATH + f'\{rit}-{tribunal}.pdf'

            tiempo_esperado = 0
            while not os.path.exists(ruta_original) and tiempo_esperado < 40:
                time.sleep(1)
                tiempo_esperado += 1

            if os.path.exists(ruta_original):                    
                try:
                    os.rename(ruta_original, ruta_caratula)
                except FileExistsError:
                    os.remove(ruta_caratula)
                    os.rename(ruta_original, ruta_caratula)
                
            else :
                print(f"error descarga documento de la causa {rit} del tribunal {tribunal}")

        for idx, row in enumerate(rows):
            if (row["Tribunal"] == tribunal) :
                excel.set_cell_value(idx+5,caratulaIndex,"OK")

        print("Descargado el tribunal "+tribunal)
        excel.save_workbook(path)

    excel.save_workbook(path)
    excel.close_workbook()

    excel.open_workbook(path) 
    rows = excel.read_worksheet_as_table("Demandas Enviadas", header=True, start=4)

    RIT_RUT = {}
        
    for row in rows:
        if row["DESCARGA_CARATULA"] == "OK" :
            if rows.columns.__contains__("4-Rut") and str(row["4-Rut"]) != 'None' :
                RIT_RUT[str(row["Rit"])+"-"+str(row["Tribunal"])] = str(row["4-Rut"])
            else :
                RIT_RUT[str(row["Rit"])+"-"+str(row["Tribunal"])] = str(row["3-Rut"])

    excel.save_workbook(path)
    excel.close_workbook()

    excel.open_workbook(EXCEL_INGRESO_DEMANDAS)
    rows = excel.read_worksheet_as_table("BOT", header=True, start=3)

    RUT_INGRESO = []
    for row in rows:
        if row["INGRESO"] == "OK" :
            RUT_INGRESO.append(str(row["ARCH_DEMANDA"]).replace('01_Demandas_firmadas\\','').split('_')[0].replace(".", "").upper())  
    
    excel.close_workbook()  
            
    merger = PdfWriter()

    for file in glob.glob(CARATULAS_FOLDER_PATH+"\*.pdf") :
        if RIT_RUT.keys().__contains__(file.replace(CARATULAS_FOLDER_PATH, '').replace("\\",'').replace(".pdf",'')) :
            if (RUT_INGRESO.__contains__(RIT_RUT[file.replace(CARATULAS_FOLDER_PATH, '').replace("\\",'').replace(".pdf",'')])) :            
                nuevo = RIT_RUT[file.replace(CARATULAS_FOLDER_PATH, '').replace("\\",'').replace(".pdf",'')] + "-" + file.replace(CARATULAS_FOLDER_PATH, '').replace("\\",'').replace(".pdf",'')
                if os.path.isfile(os.path.join(CARATULAS_FOLDER_PATH, nuevo + ".pdf")) :
                    os.remove(os.path.join(CARATULAS_FOLDER_PATH, nuevo + ".pdf"))
               
                os.rename(file, os.path.join(CARATULAS_FOLDER_PATH, nuevo + ".pdf"))
            else :
                if (file.replace(CARATULAS_FOLDER_PATH, '').replace("\\",'').replace(".pdf",'').startswith("C-")) :
                    os.remove(file)
                pass
        
    lista_ordenada = sorted(glob.glob(CARATULAS_FOLDER_PATH+"\*.pdf"))
    for file in lista_ordenada :
        merger.append(fileobj=file, pages=(0, 1))
    
    # Write to an output PDF document
    output = open(CARATULAS_FOLDER_PATH+"\CaratulasUnidas.pdf", "wb")
    merger.write(output)

    # Close file descriptors
    merger.close()
    output.close()
    print('Done Descarga Caratulas')
