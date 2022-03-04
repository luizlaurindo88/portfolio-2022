from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pyautogui as py
import FunGeral as Fg
import Usuarios as Usu
from time import sleep

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
service = Service(executable_path=r'C:\Users\01459068416\Anaconda3\chromedriver.exe')
driver = webdriver.Chrome(service=service, options=options)

link = 'https://erp.microvix.com.br/'
driver.maximize_window()
driver.get(url=link)

# variáveis

empresaBrand = "motorola"
colaBloco = 1
semaforoBrand = 0
FHora = Fg.DataFor()
dataAnterior = FHora.DataAnterior()

while semaforoBrand <= 1:

    if empresaBrand == "motorola":
        usuario = Usu.Motorola.get("usuario")
        senha = Usu.Motorola.get("senha")

    else:

        usuario = Usu.Samsung.get("usuario")
        senha = Usu.Samsung.get("senha")

        driver.find_element(By.XPATH, "//*[@id=\"logout\"]").click()
        driver.find_element(By.XPATH, "/html/body/div[11]/div/div[3]/button[1]").click()

        sleep(3)

    # loguin no portal
    driver.find_element(By.XPATH, '//*[@id="f_login"]').send_keys(usuario)
    driver.find_element(By.XPATH, '//*[@id="f_senha"]').send_keys(senha)
    sleep(1.5)
    driver.find_element(By.XPATH, '//*[@id="form_login"]/button[1]').click()
    sleep(10)

    # --selecionar a empresa
    driver.find_element(By.CLASS_NAME, "filter-option").click()
    sleep(0.5)

    driver.find_element(By.XPATH, "//*[@id='frm_selecao_empresa_login']/div/div/div/div[2]/ul/li[1]/a").click()
    driver.find_element(By.XPATH, '//*[@id="btnselecionar_empresa"]').click()
    sleep(10)

    # --aplicar os filtros
    driver.find_element(By.XPATH, '//*[@id="liModulo_0"]/a').click()
    sleep(1)
    driver.find_element(By.XPATH, '//*[@id="liModulo_0"]/ul/li[1]/a').click()
    sleep(10)

    py.PAUSE = 1.2

    # Todas as lojas
    py.click(x=1370, y=513)
    py.press('tab')
    py.press('enter')

    # ------------------------------------------------------- Aqui deve começar o loop
    if empresaBrand == "motorola":

        # clicar no campo de data
        py.doubleClick(x=490, y=935)

        py.press("delete")
        py.typewrite(dataAnterior)

        py.press("tab")

        py.press("delete")
        py.typewrite(dataAnterior)
        py.press("tab")

        sleep(1)

        # abertura do filtro "Natureza de Operação:"

        py.press(["tab", "tab"])
        py.press(["tab", "tab"])
        sleep(1)
        py.press(["tab", "tab"])
        py.press("tab")

        py.press("enter")

        sleep(2)

        # 500
        py.press(["tab", "tab"])
        py.press("enter")
        py.press("end")
        py.press("enter")

        sleep(1)

        # filtros na janela "Natureza de Operação:"
        # todos
        py.click(x=76, y=386)
        py.click(x=74, y=648)
        py.click(x=74, y=722)
        py.press('end')

        sleep(2)

        py.click(x=74, y=429)
        py.click(x=76, y=472)

        # Botao "Adicionar"
        py.press("enter")

        # filtro "Series"
        py.press(["tab", "tab"])
        py.press(["tab", "tab"])
        py.press("tab")
        py.press("enter")
        sleep(1)

        py.press(["tab", "tab"])
        py.press("enter")
        py.press("end")
        # 500:
        py.press("enter")
        sleep(1.5)

        py.click(x=75, y=384)

        py.press('end')

        sleep(1)

        py.click(x=75, y=700)
        py.press('enter')

        # enter no relatorio final
        py.press('end')

        sleep(0.5)

        py.click(x=505, y=808)

    # empresa SAMSUNG
    else:
        # Rolar a pagina
        py.doubleClick(x=1911, y=848)
        # selecionar a data

        sleep(1)
        # seleção da data
        py.doubleClick(x=481, y=258)

        py.press("delete")
        py.typewrite(dataAnterior)

        py.press("tab")

        py.press("delete")
        py.typewrite(dataAnterior)
        py.press("tab")

        sleep(1)

        # abertura do filtro "Natureza de Operação:"
        py.click(x=878, y=417)
        sleep(1)
        # Filtro "Natureza de Operação:" -> "Resultado por página"
        py.press(["tab", "tab"])
        py.press("enter")
        sleep(1)
        py.press("end")
        py.press("enter")
        sleep(2)
        # filtros na janela "Natureza de Operação:"
        # todos
        py.click(x=76, y=386)
        py.click(x=75, y=658)
        py.press('end')
        py.click(x=75, y=439)
        py.click(x=76, y=475)
        # Botao "Adicionar"
        py.press("enter")
        sleep(1)
        # filtro "Series"
        # py.click(x=881, y=517) Botao "Pesquisar"
        py.press(["tab", "tab"])
        py.press(["tab", "tab"])
        py.press("enter")
        sleep(1)

        py.press(["tab", "tab"])
        py.press("enter")
        py.press("end")
        py.press("enter")
        sleep(1)
        # inputs:
        py.click(x=76, y=385)
        py.press('end')
        sleep(1)
        py.click(x=75, y=661)
        py.press('enter')

        # gerar relatorio
        py.press('end')
        sleep(1)
        py.click(x=505, y=808),

    total = Fg.ValidaTotal()

    # neste caso serve para ordenar motorola 1º e depois a samsung para a função GeraTxt()
    semaforoBrand += 1

    empGroup = "samsungMotorola"

    Fg.GeraTxt(empresaBrand, dataAnterior, total, empGroup)

    empresaBrand = "samsung"

driver.quit()
py.alert("Fim do processo!")
