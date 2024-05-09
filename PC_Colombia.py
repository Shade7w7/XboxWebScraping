import concurrent.futures
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
import os

URL_PAGE = "https://www.xbox.com/es-co/games/all-games?cat=pcgames"

def click_button(driver):
    try:
        class_pattern = "commonStyles-module__basicButton___go-bX Button-module__basicBorderRadius___TaX9J Button-module__defaultBase___c7wIT Button-module__buttonBase___olICK Button-module__textNoUnderline___kHdUB Button-module__typeBrand___MMuct Button-module__sizeMedium___T+8s+ Button-module__overlayModeSolid___v6EcO"
        button = driver.find_element(By.CSS_SELECTOR, f"button[class*='{class_pattern}']")
        button.click()
    except Exception:
        pass

def Obtain_Data(driver):
    infobox_elements = driver.find_elements(By.CSS_SELECTOR, ".ProductCard-module__infoBox___M5x18")
    data = []

    for infobox in infobox_elements:
        title_element = infobox.find_elements(By.CSS_SELECTOR, ".ProductCard-module__title___nHGIp.typography-module__xdsBody2___RNdGY")
        price_element = infobox.find_elements(By.CSS_SELECTOR, ".Price-module__boldText___vmNHu.Price-module__moreText___q5KoT.ProductCard-module__price___cs1xr")

        if title_element and price_element:
            title = title_element[0].text
            price = price_element[0].text.replace("COP$", "").replace("+", "").strip()

            if "Ver juego" not in title:
                data.append({"title": title, "price": price})

    return data

def save_to_excel(all_data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet["A1"] = "Título"
    sheet["B1"] = "Precio"

    # Crea un diccionario para almacenar los títulos únicos y sus conteos
    unique_titles = {}

    row = 2
    for data in all_data:
        # Divide el título en partes y utiliza la parte después de 'ACA NEOGEO'
        parts = data["title"].split("ACA NEOGEO")
        title_key = (parts[-1].strip().lower() if len(parts) > 1 else data["title"].strip().lower())

        # Verifica si el título es exactamente igual a uno ya registrado
        if title_key not in unique_titles:
            sheet[f"A{row}"] = data["title"]
            sheet[f"B{row}"] = data["price"]
            row += 1
            # Añade el título en minúsculas al diccionario de títulos únicos
            unique_titles[title_key] = data["title"]

    # Obtén la ruta al perfil del usuario actual
    user_profile_path = os.environ["USERPROFILE"]

    # Construye la ruta a la carpeta de Descargas
    downloads_path = os.path.join(user_profile_path, "Downloads")

    # Especifica el nombre del archivo Excel
    excel_file_name = "Listado_juegos_pc_colombia.xlsx"

    # Combina la ruta de Descargas con el nombre del archivo para obtener la ruta completa
    file_path = os.path.join(downloads_path, excel_file_name)

    # Guarda el libro de Excel en la ruta especificada
    workbook.save(file_path)

driver = webdriver.Chrome()
driver.get(URL_PAGE)
driver.maximize_window()

all_data = []

try:
    with concurrent.futures.ThreadPoolExecutor() as executor:
        while True:
            future = executor.submit(Obtain_Data, driver)
            new_data = future.result()
            all_data.extend(new_data)
            click_button(driver)
except Exception:
    driver.quit()

save_to_excel(all_data)