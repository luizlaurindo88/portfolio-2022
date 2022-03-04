from datetime import date as dt, datetime, timedelta
import os
import pyautogui as py
from time import sleep
import clipboard as cli

imgTotais = (r"C:\\Users\\CPF\\Documents\\GitHub\\automatizacao_s_m\\totais.PNG")

# variáveis
total = ''


class Data:
    def _init_(self):
        # UTILIAZADAS NAS FUNÇÕES DataAnterior();DataAnterior2()
        self.current_date = dt.today()
        # MUDAR PARA SUBTRAIR A QUANTIDADE DE DIAS
        self.data = self.current_date - timedelta(1)
        

class DataFor (Data):
    # DATA NO FORMATO "dd/mm/yyyy"
    def DataAnterior(self):
        super()._init_()
        self.reData = dt.strftime(self.data, "%d/%m/%Y")
        return self.reData


    def DataAnterior2(self):
        super()._init_()
        self.reData2 = dt.strftime(self.data, "%d-%m-%Y")
        return self.reData2


class DataMes:

    def __init__(self, mes, ano):
        self.mesSeguinte = mes + 1
        self.mesF = dt(ano, self.mesSeguinte, 1)
        self.ultimoDia = self.mesF - timedelta(1)
        self.ultimoDia = dt.strftime(self.ultimoDia, "%d/%m/%Y")
        return self.ultimoDia


def ValidaTotal():
    sleep(20)
    py.press('end')
    sleep(1.5)
    locImigTotais = py.locateOnScreen(imgTotais)
    i = 1

    while i <= 3:
        if locImigTotais:
            py.doubleClick(locImigTotais)
            py.hotkey('ctrl', 'c')
            total = str(cli.paste())
            break

        else:
            sleep(5)
            py.press('end')
            sleep(1)
            locImigTotais = py.locateOnScreen(imgTotais)
            i += 1
    sleep(1.5)

    if total == "Totais":
        py.doubleClick(py.moveRel(800, 0))

        py.hotkey('ctrl', 'c')
        valRel = cli.paste()

        return valRel

    else:
        return False


def extMes(mes):
    
    if mes == 1: 
        exM = 'janeiro'
    elif mes == 2: 
        exM = 'fevereiro'
    elif mes == 3: 
        exM = 'março'
    elif mes == 4: 
        exM = 'abril'
    elif mes == 5: 
        exM = 'maio'
    elif mes == 6: 
        exM = 'junho'
    elif mes == 7: 
        exM = 'julho'
    elif mes == 8: 
        exM = 'agosto'
    elif mes == 9: 
        exM = 'setembro'
    elif mes == 10: 
       exM = 'outubro'
    elif mes == 11: 
       exM = 'novembro'
    else: 
       exM = 'dezembro'
    return exM


def GeraTxt(empresa, data, valor):
    
    valorDoRelatorio = str(valor)
    empData = str(empresa + "-" + data.replace("/", "-"))

    if valor is not False:
        textoEscreve = str(empresa.replace("-", ";") + " ; " + data + " ; " + valorDoRelatorio)
    else:
        textoEscreve = "Inconsistente"

    if os.path.exists(r"C:\\Users\\01459068416\\Documents\\MICROVIX\\geralLojadiario\\" + empData + ".csv"):
        with open(r"C:\\Users\\01459068416\\Documents\\MICROVIX\\geralLojadiario\\" + empData + ".csv", "a") as arquivo:
            arquivo.write("\n" + textoEscreve)
    else:
        with open(r"C:\\Users\\01459068416\\Documents\\MICROVIX\\geralLojadiario\\" + empData + ".csv", "W") as arquivo:
            arquivo.write(textoEscreve)


def GeraTxt(empresa, data, valor, empGroup):

    tempMes = datetime.strptime(data, "%d/%m/%Y").month
    tempAno = datetime.strptime(data, "%d/%m/%Y").year
    mesExtenso = str(extMes(tempMes)) + "-" + str(tempAno)

    valorDoRelatorio = str(valor)

    if empGroup != "Yes" and empGroup != "motoSamMes":
        empData = str(empGroup + "-" + data.replace("/", "-"))

    elif empGroup == "motoSamMes":

        empData = str(empGroup + "-" + mesExtenso)

    else:
        empData = str("motorolaSamsung-" + data.replace("/", "-"))

    # Montagem do texto com os dados do relatório
    if valor is not False and empGroup != "motoSamMes":
        textoEscreve = str(empresa.replace("-", ";") + " ; " + valorDoRelatorio)

    elif valor is not False and empGroup == "motoSamMes":
        textoEscreve = str(data + ";" + empresa.replace("-", ";") + " ; " + valorDoRelatorio)

    else:
        textoEscreve = "Inconsistente"

    # Caso o documento exista na pasta
    if os.path.exists(r"C:\\Users\\01459068416\\Documents\\MICROVIX\\geralLojadiario\\" + empData + ".csv"):
        with open(r"C:\\Users\\01459068416\\Documents\\MICROVIX\\geralLojadiario\\" + empData + ".csv", "a") as arquivo:
            arquivo.write("\n" + textoEscreve)
    else:
        with open(r"C:\\Users\\01459068416\\Documents\\MICROVIX\\geralLojadiario\\" + empData + ".csv", "w") as arquivo:
            arquivo.write("ORDEM; EMPRESA; VALPORTAL; VALVIEW; DIFERENÇA\n" + textoEscreve)
