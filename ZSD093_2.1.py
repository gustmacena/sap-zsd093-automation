
# DATAS PARA EXTRAÇÃO PREDEFINIDAS !

import win32com.client
from datetime import datetime
import time
import os
import tkinter as tk
from tkinter import messagebox

def close_excel():
    try:
        # Tenta obter a instância ativa do Excel e fechá-lo via COM
        excel = win32com.client.GetActiveObject("Excel.Application")
        excel.Quit()
        print("Excel fechado com sucesso via COM.")
    except Exception as e:
        print("Nenhuma instância do Excel encontrada ou erro ao fechar via COM:", e)
    try:
        # Força o encerramento do processo do Excel, se necessário
        os.system("taskkill /f /im excel.exe")
        print("Processo do Excel encerrado via taskkill.")
    except Exception as e:
        print("Erro ao encerrar o Excel via taskkill:", e)

def download_from_sap(data_inicio, data_fim):
    # Define o nome do arquivo com base nas datas informadas
    nome_arquivo = f"ZSD093 - {data_inicio} - {data_fim}.xlsx"
    try:
        # Conectar ao SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        
        # Reinicializar a transação para garantir que a tela correta seja carregada
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSD093"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        time.sleep(2)
        
        # Abrir a tela de filtro
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        time.sleep(1)
        
        # Preencher o nome do usuário e confirmar
        session.findById("wnd[1]/usr/txtENAME-LOW").text = "ANTONIOG"
        session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(1)
        
        # Selecionar a primeira linha e abrir os detalhes
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        time.sleep(1)
        
        # Preencher os filtros com as datas desejadas
        session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = data_inicio
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = data_fim
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").setFocus()
        session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").caretPosition = len(data_fim)
        time.sleep(1)
        
        # Outros filtros
        session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").text = "*"
        session.findById("wnd[0]/usr/ctxtS_VKGRP-LOW").text = "*"
        session.findById("wnd[0]/usr/ctxtS_AUART-LOW").text = "ZEPF"
        session.findById("wnd[0]/usr/ctxtS_AUART-HIGH").text = "ZEPJ"
        time.sleep(1)
        
        # Executar o relatório
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(3)
        
        # Selecionar e exportar os dados
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu()
        time.sleep(1)
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem("&XXL")
        time.sleep(2)
        
        # Confirmar a exportação e salvar o arquivo
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\gustavocm\\Desktop\\TesteE-commerce"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        time.sleep(3)
        
        # Fechar o Excel que foi aberto automaticamente
        close_excel()
        
        # Reiniciar a tela para que os controles sejam carregados novamente (F3)
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(2)
        
        print("Processo concluído com sucesso para o período:", data_inicio, "a", data_fim)
        
    except Exception as e:
        print("Erro ao executar o script:", e)

def main():
    # Lista de períodos pré-definidos (data_inicio, data_fim)
    periodos = [
        ("24.02.2025", "01.03.2025"),
        ("02.03.2025", "08.03.2025"),
        ("09.03.2025", "17.03.2025")
        # Adicione mais períodos conforme necessário
    ]
    
    # Loop para executar a extração para cada período
    for data_inicio, data_fim in periodos:
        download_from_sap(data_inicio, data_fim)
        # Aguarda alguns segundos entre as extrações para garantir que o SAP esteja pronto para o próximo ciclo
        time.sleep(5)
    
    # Ao final de todas as extrações, exibe um alerta
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    messagebox.showinfo("Concluído", "Todas as extrações definidas foram concluídas")
    root.destroy()

if __name__ == "__main__":
    main()
