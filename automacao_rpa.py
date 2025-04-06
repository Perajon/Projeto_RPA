import pandas as pd
import pyautogui
import time
from openpyxl import Workbook
from posicoes import posicoes

def clicar(dado):
    if dado in posicoes:
        x, y = posicoes[dado]
        pyautogui.click(x, y)
        return True
    return False

def digitar(texto):
    pyautogui.write(texto)
    return True

def pressionar(tecla):
    pyautogui.press(tecla)
    return True

def esperar(segundos):
    try:
        segundos = int(segundos)
        time.sleep(segundos)
        return True
    except:
        return False

def executar_tarefas(arquivo):
    df = pd.read_csv(arquivo)
    relatorio = []

    for index, row in df.iterrows():
        tarefa = row['Tarefa']
        tipo = row['Tipo']
        dado = str(row['Dado'])
        status = False

        print(f"Executando: {tarefa} ({tipo})...")

        if tipo == 'click':
            status = clicar(dado)
        elif tipo == 'texto':
            status = digitar(dado)
        elif tipo == 'tecla':
            status = pressionar(dado)
        elif tipo == 'espera':
            status = esperar(dado)

        tempo_exec = time.strftime("%H:%M:%S", time.localtime())
        relatorio.append([tarefa, tipo, dado, "Sucesso" if status else "Falha", tempo_exec])
        time.sleep(1)

    return relatorio

def gerar_relatorio(relatorio, nome_arquivo="relatorio.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.append(["Tarefa", "Tipo", "Dado", "Status", "Horário"])

    for linha in relatorio:
        ws.append(linha)

    wb.save(nome_arquivo)
    print(f"Relatório salvo como {nome_arquivo}")

if __name__ == "__main__":
    relatorio = executar_tarefas("tarefas.csv")
    gerar_relatorio(relatorio)
