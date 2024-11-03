# Libs
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service


# Functions
def read_excel_file(file_path):
    '''
    Read Excel File using Pandas
    
    Parameters
    ----------
    file_path : str
        Excel file path to be read
        
    Returns
    ----------
    df: pandas.dataframe
        Pandas dataframe with Excel file data on it
    '''
    df = pd.read_excel(file_path)
    return df


# Get data from Excel file / Obter dados da planilha Excel
excelFilePath = r'.\challenge.xlsx'
excelDf = read_excel_file(excelFilePath)

# Config ChromeDriver options
opts = Options()
opts.add_experimental_option("useAutomationExtension", False)
opts.add_experimental_option("excludeSwitches",["enable-automation"])
opts.add_experimental_option("detach", True)

# Config Webdriver
service = Service(ChromeDriverManager().install())

# Init driver with custom options
driver = webdriver.Chrome(service=service, options=opts)

# Url RPA Challenge
url = 'https://rpachallenge.com/'

# Go to URL
driver.get(url)

# Loop for each row in dataframe (data from Excel file)
for index, row in excelDf.iterrows():
    # Check if is the first row to click on Start button
    if index == 0:
        # Click start button
        startButton = driver.find_element(By.XPATH, '//button[text()="Start"]')
        startButton.click()

    # Set variables
    first_name = row['First Name']
    last_name = row['Last Name ']
    company_name = row['Company Name']
    role = row['Role in Company']
    address = row['Address']
    email = row['Email']
    phone_number = row['Phone Number']

    # Set value of the First Name input
    firstNameInput = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelFirstName"]')
    firstNameInput.send_keys(first_name)
    
    # Set value of the Last Name input
    lastNameInput = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelLastName"]')
    lastNameInput.send_keys(last_name)
    
    # Set value of the Company Name input
    companyNameInput = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelCompanyName"]')
    companyNameInput.send_keys(company_name)

    # Set value of the Role in Company input
    roleInput = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelRole"]')
    roleInput.send_keys(role)

    # Set value of the Address input
    addressInput = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelAddress"]')
    addressInput.send_keys(address)

    # Set value of the Email input
    emailInput = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelEmail"]')
    emailInput.send_keys(address)

    # Set value of the Email input
    phoneNumberInput = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelPhone"]')
    phoneNumberInput.send_keys(phone_number)

    # Click Submit button
    submitButton = driver.find_element(By.XPATH, '//input[@type="submit"]')
    submitButton.click()

# Take a screenshot of the screen
driver.save_screenshot("screenshots/challenge-completed.png")

# Close driver
driver.quit()






