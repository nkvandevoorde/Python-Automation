#Packages: Selenium, pansas, unicodedata, pathlib, and openpyxl (installed not imported)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import unicodedata 
from pathlib import Path

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
driver.get("https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp?search=9437495")
print(driver.title) #check for correct website

#Locate search bar
search_box = driver.find_element(by = By.ID, value = "pesquisaValue")

#Input CNES values from cnes_df into search bar and submit
cnes_grabs = []
for cnes in cnes_entries:
    search_box.clear()  #Clear the search box
    search_box.send_keys(cnes)  #Enter the CNES value
    search_box.send_keys(Keys.RETURN)  #Submit
    driver.implicitly_wait(3)  #Wait for the page to load
    #Grab and store address and coordinates for each CNES value
    plus_button = driver.find_element(
        By.XPATH,
        "//button[@ng-click='buscarEstabalecimento(estab.id)']"
    )
    if plus_button:
        plus_button.click()  #Click the plus button to view details 
        address_bar = driver.find_element(
            By.XPATH("//input[@id='cnpj' and @ng-value='estabelecimento.noLogradouro']")
        )
        address = address_bar.text.strip()
        cnes_grabs.append((cnes, address))
    else:
        cnes_grabs.append((cnes, "No address found"))
print(len(cnes_grabs)) #check for correct number of grabs

#Input CNES values from relatorio_df into search bar and submit
relatorio_grabs = []
for cnes in relatorio_entries:
    search_box.clear()  #Clear the search box
    search_box.send_keys(cnes)  #Enter the CNES value
    search_box.send_keys(Keys.RETURN)  #Submit
    driver.implicitly_wait(3)  #Wait for the page to load
    #Grab and store address and coordinates for each CNES value
    plus_button = driver.find_element(
        By.XPATH,
        "//button[@ng-click='buscarEstabalecimento(estab.id)']"
    )
    if plus_button:
        plus_button.click()  #Click the plus button to view details 
        address_bar = driver.find_element(by = By.XPATH("//input[@id='cnpj' and @ng-value='estabelecimento.noLogradouro']"))
        address = address_bar.text.strip()
        cnes_grabs.append((cnes, address))
    else:
        cnes_grabs.append((cnes, "No address found"))
print(len(relatorio_grabs)) #check for correct number of grabs

driver.quit() #close driver

#write cnes_grabs to excel file
cnes_output_df = pd.DataFrame(cnes_grabs, columns=['CNES', 'Address'])
cnes_output_file_path = script_dir / "cnes_output.xlsx"
cnes_output_df.to_excel(cnes_output_file_path, index=False)

#write relatorio_grabs to excel file
relatorio_output_df = pd.DataFrame(relatorio_grabs, columns=['CNES', 'Address'])
relatorio_output_file_path = script_dir / "relatorio_output.xlsx"
relatorio_output_df.to_excel(relatorio_output_file_path, index=False)

