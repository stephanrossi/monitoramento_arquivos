import time
import threading
import pandas as pd
import os
import sys
import win32file
import win32con
import win32security
import ntsecuritycon as con
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pystray
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import filedialog
import logging

# Configurar o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Variáveis globais
MONITORED_DIR = None
REPORT_NAME = None
event_list = []
observer = None
icon = None
event_list_lock = threading.Lock()

# Função para obter o nome do usuário que modificou o arquivo
def get_file_owner(filepath):
    sd = win32security.GetFileSecurity(filepath, win32security.OWNER_SECURITY_INFORMATION)
    owner_sid = sd.GetSecurityDescriptorOwner()
    name, domain, type = win32security.LookupAccountSid(None, owner_sid)
    return f"{domain}\\{name}"

# Classe para lidar com eventos do sistema de arquivos
class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            try:
                filepath = event.src_path
                filename = os.path.basename(filepath)
                file_size_bytes = os.path.getsize(filepath)
                file_size_kb = file_size_bytes / 1024  # Convertendo para KB
                file_size_kb = round(file_size_kb, 2)  # Arredondando para duas casas decimais
                file_owner = get_file_owner(filepath)
                timestamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                event_info = {
                    'Data/Hora': timestamp,
                    'Usuário': file_owner,
                    'Evento': 'Criado',
                    'Arquivo': filename,
                    'Caminho': filepath,
                    'Tamanho (KB)': file_size_kb
                }
                with event_list_lock:
                    event_list.append(event_info)
                logging.info(f"Arquivo criado: {filename}")
            except Exception as e:
                logging.error(f"Erro ao processar o arquivo {event.src_path}: {e}")

    def on_modified(self, event):
        if not event.is_directory:
            try:
                filepath = event.src_path
                filename = os.path.basename(filepath)
                file_size_bytes = os.path.getsize(filepath)
                file_size_kb = file_size_bytes / 1024  # Convertendo para KB
                file_size_kb = round(file_size_kb, 2)  # Arredondando para duas casas decimais
                file_owner = get_file_owner(filepath)
                timestamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                event_info = {
                    'Data/Hora': timestamp,
                    'Usuário': file_owner,
                    'Evento': 'Modificado',
                    'Arquivo': filename,
                    'Caminho': filepath,
                    'Tamanho (KB)': file_size_kb
                }
                with event_list_lock:
                    event_list.append(event_info)
                logging.info(f"Arquivo modificado: {filename}")
            except Exception as e:
                logging.error(f"Erro ao processar o arquivo {event.src_path}: {e}")

    def on_deleted(self, event):
        if not event.is_directory:
            try:
                filepath = event.src_path
                filename = os.path.basename(filepath)
                file_owner = "Desconhecido"
                timestamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                event_info = {
                    'Data/Hora': timestamp,
                    'Usuário': file_owner,
                    'Evento': 'Deletado',
                    'Arquivo': filename,
                    'Caminho': filepath,
                    'Tamanho (KB)': 'N/A'
                }
                with event_list_lock:
                    event_list.append(event_info)
                logging.info(f"Arquivo deletado: {filename}")
            except Exception as e:
                logging.error(f"Erro ao processar o arquivo {event.src_path}: {e}")

# Função para gerar o relatório em Excel a cada 5 minutos
def schedule_report():
    while True:
        time.sleep(300)  # Espera 5 minutos (300 segundos)
        with event_list_lock:
            if event_list and REPORT_NAME:
                events_to_process = event_list.copy()
                event_list.clear()
            else:
                events_to_process = []

        if events_to_process:
            logging.info("Salvando novos eventos no relatório.")
            try:
                if os.path.exists(REPORT_NAME):
                    existing_df = pd.read_excel(REPORT_NAME)
                    new_df = pd.DataFrame(events_to_process)
                    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                else:
                    combined_df = pd.DataFrame(events_to_process)
                combined_df.to_excel(REPORT_NAME, index=False)
            except Exception as e:
                logging.error(f"Erro ao salvar o relatório: {e}")
        else:
            logging.info("Nenhum novo evento para salvar.")

# Funções para o menu do ícone da bandeja
def on_exit(icon, item):
    icon.stop()
    if observer:
        observer.stop()
        observer.join()
    # Salva os eventos restantes
    with event_list_lock:
        if event_list and REPORT_NAME:
            logging.info("Salvando eventos restantes antes de sair.")
            try:
                if os.path.exists(REPORT_NAME):
                    existing_df = pd.read_excel(REPORT_NAME)
                    new_df = pd.DataFrame(event_list)
                    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                else:
                    combined_df = pd.DataFrame(event_list)
                combined_df.to_excel(REPORT_NAME, index=False)
                event_list.clear()
            except Exception as e:
                logging.error(f"Erro ao salvar o relatório: {e}")

def select_monitored_dir(icon, item):
    global MONITORED_DIR, observer
    root = tk.Tk()
    root.withdraw()
    new_dir = filedialog.askdirectory(title="Selecione o diretório a ser monitorado")
    if new_dir:
        MONITORED_DIR = new_dir
        logging.info(f"Diretório monitorado atualizado para: {MONITORED_DIR}")
        # Reiniciar o observador se já estiver rodando
        if observer:
            observer.stop()
            observer.join()
            event_handler = MyHandler()
            observer = Observer()
            observer.schedule(event_handler, path=MONITORED_DIR, recursive=True)
            observer.start()

def select_report_location(icon, item):
    global REPORT_NAME
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        title="Selecione onde salvar o relatório",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        REPORT_NAME = file_path
        logging.info(f"Local do relatório definido para: {REPORT_NAME}")

def create_image():
    # Verifica se está sendo executado pelo PyInstaller
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    icon_path = os.path.join(base_path, "folder.ico")
    return Image.open(icon_path)

def setup_tray_icon():
    global icon
    menu = pystray.Menu(
        pystray.MenuItem("Selecionar diretório monitorado", select_monitored_dir),
        pystray.MenuItem("Selecionar local para salvar o relatório", select_report_location),
        pystray.MenuItem("Sair", on_exit)
    )
    image = create_image()
    icon = pystray.Icon("Monitor de arquivos", image, "Monitor de arquivos", menu)
    icon.run()

def main():
    global observer
    # Solicita ao usuário o diretório e o local do relatório
    select_monitored_dir(None, None)
    select_report_location(None, None)
    if not MONITORED_DIR or not REPORT_NAME:
        logging.error("Diretório monitorado ou local do relatório não foi selecionado. Encerrando.")
        return

    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path=MONITORED_DIR, recursive=True)
    observer.start()

    # Thread para agendar o relatório
    report_thread = threading.Thread(target=schedule_report)
    report_thread.daemon = True
    report_thread.start()

    # Configurar o ícone da bandeja
    setup_tray_icon()

    # Parar o observador quando o ícone da bandeja for encerrado
    observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
