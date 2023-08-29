from selenium import webdriver
import openpyxl
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service


def main():
    # Instala o pacote selenium
    servico = Service(ChromeDriverManager().install())
    # Seta o caminho do arquivo
    path_xlsx = r'C:/Users/Joaoc/Downloads/challenge.xlsx'
    # Lê Planilha
    wkb = openpyxl.load_workbook((path_xlsx))
    sheet = wkb['Sheet1']
    navegador = webdriver.Chrome(service=servico)

    navegador.get("https://rpachallenge.com/")
    #Seta botão de start
    xpath_botao = '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button'
    #Clica no botão de start
    navegador.find_element('xpath',
                       xpath_botao
                       ).click()
    #for que insere cada campo em cada campo
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        try:

            navegador.find_element('xpath',
                               '//*[@ng-reflect-name="labelEmail"]'
                               ).send_keys(row[5])

            navegador.find_element('xpath',
                               '//*[@ng-reflect-name="labelPhone"]'
                               ).send_keys(row[6])

            navegador.find_element('xpath',
                               '//*[@ng-reflect-name="labelFirstName"]'
                               ).send_keys(row[0])

            navegador.find_element('xpath',
                               '//*[@ng-reflect-name="labelLastName"]'
                               ).send_keys(row[1])

            navegador.find_element('xpath',
                               '//*[@ng-reflect-name="labelCompanyName"]'
                               ).send_keys(row[2])

            navegador.find_element('xpath',
                               '//*[@ng-reflect-name="labelRole"]'
                               ).send_keys(row[3])

            navegador.find_element('xpath',
                                '//*[@ng-reflect-name="labelAddress"]'
                                ).send_keys(row[4])

            xpath_b = '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input'

            navegador.find_element('xpath',
                               xpath_b
                               ).click()
        except Exception as e:
            print(e)

    input("insira o dado")

    wkb.close()

if __name__ == "__main__":
    main()