## Packages: Selenium, pansas, unicodedata, pathlib, and openpyxl (installed not imported)
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import unicodedata  
from pathlib import Path

## Helpers
def find_search_box (driver):
    return driver.find_element(by = By.ID, value = "pesquisaValue")

def wait_for_results(driver):
    WebDriverWait(driver, 10).until(lambda d: len(d.find_elements(By.XPATH, "//button[@ng-click='buscarEstabalecimento(estab.id)']")) >= 0 )

def click(driver, xpath, retries=3):
    for _ in range(retries):
        try:
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            return True
        except StaleElementReferenceException:
            continue    
    return False

def get_address (driver, xpath):
    try:
        return WebDriverWait(driver, 10).until(lambda d: next((el.get_attribute("value").strip()
                for el in d.find_elements(By.XPATH, xpath)
                    if el.is_displayed() and el.get_attribute("value").strip()), None))
    except TimeoutException:
            return None 

#Remove special characters for consistent spelling
def remove_accents(input_str):
    if input_str is None:
        return None
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd_form if not unicodedata.combining(c)])

#Find, Load, and Clean CNES data file
script_dir = Path(__file__).parent.absolute()
file_path = script_dir / "CNES_2024_with_REGIC_labels_p.xlsx"
cnes_df = pd.read_excel(file_path)
cnes_df = cnes_df.drop_duplicates()

#Find, load, and clean Relatorio file
relatorio_file_path = script_dir / "relatorio-geral_COPY0510.xlsx"
relatorio_df = pd.read_excel(relatorio_file_path)
relatorio_df = relatorio_df.drop_duplicates()
relatorio_df = relatorio_df.dropna(subset=['Unnamed: 12'])
#print(relatorio_df.columns) #check for correct columns

#Create list of CNES values from cnes_df
cnes_entries = []
for index, row in cnes_df.iterrows():
    city_val = row['CNES']  # Column 'F' (CNES)
    
    if all(v is not None for v in [city_val]):
        cnes_entries.append((
            remove_accents(str(city_val).strip().lower())
        ))
print(len(cnes_entries)) #check for correct number of entries

#Create list of CNES values from relatorio_df
relatorio_entries = []
for index, row in relatorio_df.iterrows():
    city_val = row['Unnamed: 12']  # Column 'M' (CNES)
    
    if all(v is not None for v in [city_val]):
        relatorio_entries.append((
            remove_accents(str(city_val).strip().lower())
        ))
print(len(relatorio_entries)) #check for correct number of entries

#Open driver and bypass privacy error
options = webdriver.ChromeOptions()
options.add_argument('--ignore-ssl-errors=yes')
options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(options=options)

#Navigate to CNES website
driver.get("https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp?")
print(driver.title) #check for correct website

# #Input CNES values from cnes_df into search bar and submit
# cnes_grabs = []

# for cnes in cnes_entries:
#     try:
#         search_box = find_search_box(driver)  #Find search box
#         search_box.clear()  #Clear the search box
#         search_box.send_keys(cnes, Keys.RETURN)  #Submit
#         wait_for_results(driver)  #Wait for results to load
        
#         plus_button_xpath = "//button[@ng-click='buscarEstabalecimento(estab.id)']"
#         if click(driver, plus_button_xpath):

#             # wait for modal to appear
#             modal = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located(
#             (By.XPATH, "//div[contains(@class,'modal')]")
#                 )
#             )

#             # find visible address input inside modal
#             address = get_address(driver, "//div[contains(@class,'modal')]//input[@ng-value='estabelecimento.noLogradouro']")
#             number = get_address(driver, "//div[contains(@class,'modal')]//input[@ng-value= 'estabelecimento.nuEndereco']")
#             zipCode = get_address(driver, "//div[contains(@class,'modal')]//input[@ui-mask='99999-999']")
            
#             #append address and check to make sure the program is finding real addresses
#             cnes_grabs.append((cnes, number if number else "No number found", address if address else "No address found", zipCode if zipCode else "No zip code found"))
#             if len(cnes_grabs) % 100 == 0:
#                 print(
#                         f"[CHECKPOINT] Processed {len(cnes_grabs)} records | "
#                         f"Last CNES: {cnes} | "
#                         f"Last number: {repr(number)} | "
#                         f"Last address: {repr(address)} | "
#                         f"Last zip code: {repr(zipCode)}"
#                 )
#             # close modal
#             click(driver, "//button[contains(@class,'close')]")
#         else:
#             cnes_grabs.append((cnes, "No address found"))
#     except StaleElementReferenceException:
#         continue
#     except TimeoutException:
#         cnes_grabs.append((cnes, "No address found / Timeout"))

# #write cnes_grabs to excel file
# cnes_output_df = pd.DataFrame(cnes_grabs, columns=['CNES', 'Number', 'Address', 'Zip Code'])
# cnes_output_file_path = script_dir / "cnes_output.xlsx"
# cnes_output_df.to_excel(cnes_output_file_path, index=False)

#Input CNES values from cnes_df into search bar and submit
relatorio_grabs = []

for cnes in relatorio_entries:
    try:
        search_box = find_search_box(driver)  #Find search box
        search_box.clear()  #Clear the search box
        search_box.send_keys(cnes, Keys.RETURN)  #Submit
        wait_for_results(driver)  #Wait for results to load
        
        plus_button_xpath = "//button[@ng-click='buscarEstabalecimento(estab.id)']"
        if click(driver, plus_button_xpath):
                        #wait for modal to appear
            modal = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
            (By.XPATH, "//div[contains(@class,'modal')]")
                )
            )

            #find visible address input inside modal
            address = get_address(driver, "//div[contains(@class,'modal')]//input[@ng-value='estabelecimento.noLogradouro']")
            number = get_address(driver, "//div[contains(@class,'modal')]//input[@ng-value= 'estabelecimento.nuEndereco']")
            zipCode = get_address(driver, "//div[contains(@class,'modal')]//input[@ui-mask='99999-999']")

            #append address and check to make sure the program is finding real addresses
            relatorio_grabs.append((cnes, number if number else "No number found", address if address else "No address found", zipCode if zipCode else "No zip code found"))
            if len(relatorio_grabs) % 100 == 0:
                print(
                        f"[CHECKPOINT] Processed {len(relatorio_grabs)} records | "
                        f"Last CNES: {cnes} | "
                        f"Last number: {repr(number)} | "
                        f"Last address: {repr(address)} | "
                        f"Last zip code: {repr(zipCode)}"
                )
            #close modal
            click(driver, "//button[contains(@class,'close')]")
        else:
            relatorio_grabs.append((cnes, "No address found"))
    except StaleElementReferenceException:
        continue
    except TimeoutException:
        relatorio_grabs.append((cnes, "No address found / Timeout"))

#write relatorio_grabs to excel file
relatorio_output_df = pd.DataFrame(relatorio_grabs, columns=['CNES', 'Number', 'Address', 'Zip Code'])
relatorio_output_file_path = script_dir / "relatorio_output.xlsx"
relatorio_output_df.to_excel(relatorio_output_file_path, index=False)