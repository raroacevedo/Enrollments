import re
import pandas as pd
import os, gc   
from time import sleep
from getpass import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

#configura el driver de selenium
def setup_driver():
    """Configura y retorna el WebDriver con opciones seguras"""
    service = Service(r"..\Chrome\chromedriver.exe")

    # Verifica si el chromedriver existe en la ruta especificada
    if not os.path.exists(service.path):    
        raise FileNotFoundError(f"El chromedriver no se encuentra en la ruta: {service.path}")
    
    # Configuración de opciones del navegador
    options = Options()
    options.add_argument("--disable-extensions")
    # options.add_argument("--headless")  # Descomentar si se desea en modo headless
    options.add_argument("--log-level=3")
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(10)
    return driver

#login con credenciales y clave 2FA
def login(driver, username, password, second_factor):
    """Realiza login con credenciales y clave 2FA"""
    driver.get("https://virtual.upb.edu.co/d2l/login?noRedirect=1")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "userName"))).send_keys(username)
    #WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password"))).send_keys(password + Keys.RETURN)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password"))).send_keys("2/2*2Hijosmemsam" + Keys.RETURN)
    sleep(1)

    # Autenticación de dos factores
    driver.get("https://virtual.upb.edu.co/d2l/lp/auth/twofactorauthentication/TwoFactorCodeEntry.d2l")
    #WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "z_d"))).send_keys(second_factor + Keys.RETURN) //codigo anterior
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "z_i"))).send_keys(second_factor + Keys.RETURN)

    # Esperar hasta que se redireccione al home de la plataforma para asegurar que el login se complete
    sleep(3)

#extrae el ID del curso desde la URL usando expresiones regulares
def get_course_id(url):
    """Extrae el ID del curso (longitud: entre 4 y 6) desde la URL usando expresiones regulares"""
    match = re.search(r'/(\d{4,6})$', url.strip())
    return match.group(1) if match else None
       
#extrae el shortname del curso 
def get_shortname(driver, course_id, error_log):
    """Navega a la página del curso y extrae el shortname si está disponible"""
    url = f'https://virtual.upb.edu.co/d2l/lp/manageCourses/course_offering_info_viewedit.d2l?ou={course_id}'
    try:
        driver.get(url)
        sleep(1)  # Espera para que la página cargue completamente

        shortname_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "z_l"))
        )
        sleep(1)  # Espera para que la página cargue completamente
        return shortname_field.get_attribute('value')

    except Exception as e:
        print(f"[Error] No se pudo acceder o encontrar el campo Código de oferta de curso para el ID {course_id}")
        error_log.append({
            'course_id': course_id,
            'url': url,
            'error': str(e)
        })
        return None

#valida si el curso ya existe en el archivo shortnames.csv
def load_existing_shortnames(file_path):
    """Carga datos previos si el archivo existe, para evitar duplicados"""
    if os.path.exists(file_path):
        return pd.read_csv(file_path)
    return pd.DataFrame(columns=['Nombre', 'NRC'])


def main():
    # --- Recolección de credenciales ---
    user = input('Username: ')
    password = getpass('Password: ')
    second_factor = input("Clave 2FA: ")

    # --- Configuración del navegador ---
    driver = setup_driver()

    try:
        login(driver, user, password, second_factor)
 
        # --- Lectura de URLs desde CSV ---
        df_urls = pd.read_csv('ListaCursos.csv')
        course_urls = df_urls['Enlace curso'].dropna() # Asegurarse de que no haya valores nulos

        output_file = 'shortnames.csv'
        existing_data = load_existing_shortnames(output_file)
        existing_names = set(existing_data['Nombre'])

        new_course_data = []
        errores = []  # Lista de errores

        for url in course_urls:
            course_id = get_course_id(url)
            if not course_id:
                print(f"[Advertencia] No se pudo extraer ID del curso desde URL: {url}")
                errores.append({
                    'course_id': 'N/A',
                    'url': url,
                    'error': 'No se pudo extraer el ID del curso'
                })
                continue

            print(f"\nProcesando curso ID: {course_id}")
            short_name = get_shortname(driver, course_id, errores)

            if short_name and short_name not in existing_names:
                #Se extrae el periodo y NRC/LC
                partes = short_name.split('-')
                if len(partes) == 4:
                    Periodo = partes[-2]
                    NRC_LC = partes[-1]
                    
                    new_course_data.append((short_name, NRC_LC, Periodo))
                    existing_names.add(short_name)
                else:
                    print(f"[Advertencia] EL codigo del curso tiene errores: {short_name}")

            elif short_name:
                print(f"[Info] Curso ya registrado previamente: {short_name}")

        # Guardar resultados exitosos
        if new_course_data:
            df_new = pd.DataFrame(new_course_data, columns=['Nombre', 'NRC', 'Periodo'])
            df_result = pd.concat([existing_data, df_new], ignore_index=True)
            df_result.to_csv(output_file, index=False)
            print(f"\n[✓] Datos nuevos agregados. Archivo actualizado: {output_file}")
        else:
            print("\n[✓] No se encontraron nuevos cursos para agregar.")

        # Guardar errores
        if errores:
            df_errores = pd.DataFrame(errores)
            df_errores.to_csv('errores.csv', index=False)
            print(f"\n[!] Se encontraron errores. Ver detalle en errores.csv")

    finally:
        driver.quit()
        del driver, df_result, df_new            # Limpieza de variables   
        gc.collect()                             # Limpieza de memoria
        print("\n[✓] Proceso finalizado.")

if __name__ == '__main__':
    main()