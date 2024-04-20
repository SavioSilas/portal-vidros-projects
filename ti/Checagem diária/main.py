import os, cv2, ssl, sys, imaplib, email, threading
import datetime, requests, urllib3
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import numpy as np
import pandas as pd
from email.header import decode_header
import tkinter as tk
from tkinter import messagebox, Toplevel, Label, Entry, Button
from pathlib import Path

def finish_process():
    status_label.config(text="Processo Finalizado!")
    messagebox.showinfo("svosvosvo")

# Janela principal após o login
def main_window():
    global root, status_label
    root = tk.Tk()
    root.title("Checagem Diária")
    root.geometry("500x300")
    process_button = Button(root, text="Iniciar Processo", command=start_process)
    process_button.pack(pady=20)

    # Label para mostrar o status do processo
    status_label = Label(root, text="")
    status_label.pack(pady=10)

    root.mainloop()

def login():
    username = entry_username.get()
    password = entry_password.get()
    if username == 's' and password == 's':  
        login_window.destroy()
        main_window()
    else:
        messagebox.showerror("Erro", "Usuário ou senha incorretos!")

# Função para iniciar o processo principal em uma thread separada
def start_process():
    status_label.config(text="Processo em Andamento...")
    threading.Thread(target=main).start()

def verificar_backup_webglass(diretorio):
    print("1 VERIFICAR BACKUP DO WEBGLASS")
    meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    mes_atual = meses[datetime.datetime.now().month - 1]

    if not os.path.exists(diretorio):
        print("· Diretório não encontrado.")
        return

    for nome in os.listdir(diretorio):
        if nome.startswith(mes_atual):
            caminho_completo = os.path.join(diretorio, nome)
            if os.path.isdir(caminho_completo):
                timestamp = os.path.getctime(caminho_completo)
                data_criacao = datetime.datetime.fromtimestamp(timestamp)
                print(f"· Nome da pasta: {nome}")
                print(f"· Data de criação: {data_criacao.strftime('%d/%m/%Y')}")
                print(f"· Hora de criação: {data_criacao.strftime('%H:%M:%S')}")

                print("Subpastas:")
                subpastas = []
                for subnome in os.listdir(caminho_completo):
                    subcaminho = os.path.join(caminho_completo, subnome)
                    if os.path.isdir(subcaminho):
                        sub_timestamp = os.path.getctime(subcaminho)
                        subpastas.append((subnome, sub_timestamp))

                # Ordenar subpastas por data de criação e pegar as duas últimas
                subpastas.sort(key=lambda x: x[1], reverse=True)
                for subnome, sub_timestamp in subpastas[:2]:
                    sub_data_criacao = datetime.datetime.fromtimestamp(sub_timestamp)
                    print(f"· Nome da subpasta: {subnome}")
                    print(f"· Data de criação: {sub_data_criacao.strftime('%d/%m/%Y')}")
                    print(f"· Hora de criação: {sub_data_criacao.strftime('%H:%M:%S')}")
                return
    print("· Pasta do mês atual não encontrada.")

def verificar_site(url):
    print("\n2 VERIFICAR O SITE DO WEBGLASS")
    try:
        resposta = requests.get(url, timeout=10)  
        if resposta.status_code == 200:
            print(f"· O site {url} está no ar!\n")
        else:
            print(f"· O site {url} está acessível, mas retornou um código de status {resposta.status_code}.\n")
    except requests.exceptions.RequestException as e:
        print(f"· Não foi possível acessar o site {url}. Erro: {e}")

def verificar_backup_fortes(diretorio):
    print("3 VERIFICAR BACKUP DO FORTES")

    # Verifica se o diretório existe
    if not os.path.exists(diretorio):
        print("· Diretório não encontrado.")
        return

    # Lista todas as pastas dentro do diretório especificado
    pastas = [d for d in os.listdir(diretorio) if os.path.isdir(os.path.join(diretorio, d))]
    if not pastas:
        print("· Nenhuma pasta encontrada no diretório.")
        return

    # Itera sobre cada pasta encontrada e imprime suas informações
    for nome in pastas:
        caminho_completo = os.path.join(diretorio, nome)
        timestamp = os.path.getmtime(caminho_completo)
        data_criacao = datetime.datetime.fromtimestamp(timestamp)
        print(f"· Nome da pasta: {nome}")
        print(f"· Data de criação: {data_criacao.strftime('%d/%m/%Y')}")
        print(f"· Hora de criação: {data_criacao.strftime('%H:%M:%S')}")

def verificar_vms():
    print("\n5 VERIFICAR VM's")
    # Configurações do servidor ESXi
    host = "s"
    user = "s"
    password = 'svo'
    
    # Conecta ao servidor ESXi
    context = ssl._create_unverified_context()
    si = SmartConnect(host=host, user=user, pwd=password, sslContext=context)
    
    # Obter o objeto de content do vSphere
    content = si.RetrieveContent()
    
    # Obter todas as VMs
    container = content.viewManager.CreateContainerView(content.rootFolder, [vim.VirtualMachine], True)
    vms = container.view
    container.Destroy()

    # Verificar o status de cada VM
    for vm in vms:
        print(f"· VM: {vm.name} - Power State: {vm.runtime.powerState}")

    # Desconectar do servidor
    Disconnect(si)

# Suprime os avisos de requisições HTTPS sem verificação de SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
def verificar_controladoras(matriz, filial):
    print("\n6 VERIFICAR CONTROLADORA UNIFI MATRIZ/FILIAL")
    print("MATRIZ")
    try:
        response = requests.get(matriz, verify=False, timeout=10)
        if response.status_code == 200:
            print(f"· O site {matriz} está no ar!")
        else:
            print(f"· O site {matriz} está no ar, mas retornou um código de status {response.status_code}.")
    except requests.exceptions.ConnectionError:
        print(f"· Não foi possível conectar ao site {matriz}.")
    except requests.exceptions.Timeout:
        print(f"· O pedido ao site {matriz} excedeu o tempo limite de conexão.")
    except requests.exceptions.RequestException as e:
        print(f"· Erro ao acessar o site {matriz}: {e}")

    print("FILIAL")
    try:
        response = requests.get(filial, verify=False, timeout=10)
        if response.status_code == 200:
            print(f"· O site {filial} está no ar!")
        else:
            print(f"· O site {filial} está no ar, mas retornou um código de status {response.status_code}.")
    except requests.exceptions.ConnectionError:
        print(f"· Não foi possível conectar ao site {filial}.")
    except requests.exceptions.Timeout:
        print(f"· O pedido ao site {filial} excedeu o tempo limite de conexão.")
    except requests.exceptions.RequestException as e:
        print(f"· Erro ao acessar o site {filial}: {e}")

def check_camera(url, channel, crop_size=(224, 224)):
    cap = cv2.VideoCapture(url)
    if not cap.isOpened():
        return False, f"Não foi possível acessar a câmera {channel}"

    ret, frame = cap.read()
    if not ret:
        cap.release()
        return False, f"Não foi possível capturar imagem da câmera {channel}"

    start_x = (frame.shape[1] - crop_size[0]) // 2
    start_y = (frame.shape[0] - crop_size[1]) // 2
    end_x = start_x + crop_size[0] + 2500
    end_y = start_y + crop_size[1] + 1800

    cropped_frame = frame[start_y:end_y, start_x:end_x]
    cap.release()

    if np.all(cropped_frame == 0):
        return False, f"A imagem da câmera {channel} está com a imagem preta\n"
    
    return True, None

def checar_cameras():
    cameras_externo = []
    cameras_adm = []
    cameras_producao = []
    cameras_filial1 = []
    cameras_filial2 = []

    print("\n7 VERIFICAR CAMERAS")
    print("CÂMERAS DA MATRIZ - s EXTERNO/COMERCIAL")
    for channel in range(1, 17):
        url_matriz_18 = f"s"
        ok, problema = check_camera(url_matriz_18, channel)
        if not ok:
            cameras_externo.append(problema)
    
    if cameras_externo:
        print("· Todas as câmeras estão ok. \n Exceto a(s) câmera(s)", ', '.join(cameras_externo))
    else:
        print("· Todas as câmeras da MATRIZ - EXTERNO/COMERCIAL estão ok")

    print("CÂMERAS DA MATRIZ - s ADMINISTRAÇÃO")
    for channel in range(1, 17):
        url_matriz_19 = f"s"
        ok, problema = check_camera(url_matriz_19, channel)
        if not ok:
            cameras_adm.append(problema)
    
    if cameras_adm:
        print("· Todas as câmeras estão ok. \n Exceto a(s) câmera(s)", ', ', '\t'.join(cameras_adm))
    else:
        print("· Todas as câmeras DA MATRIZ - ADMINISTRAÇÃO estão ok")

    print("CÂMERAS DA MATRIZ - s PRODUÇÃO")
    for channel in range(1, 17):
        url_matriz_20 = f"iavso"
        ok, problema = check_camera(url_matriz_20, channel)
        if not ok:
            cameras_producao.append(problema)
    
    if cameras_producao:
        print("· Todas as câmeras estão ok. \n Exceto a(s) câmera(s)", ', '.join(cameras_producao))
    else:
        print("· Todas as câmeras da MATRIZ - PRODUÇÃO estão ok")
    print("CÂMERAS DA FILIAL - s EXTERNO/COMERCIAL")
    for channel in range(1, 17):
        url_filial_133 = f"s"
        ok, problema = check_camera(url_filial_133, channel)
        if not ok:
            cameras_filial1.append(problema)
    
    if cameras_filial1:
        print("· Todas as câmeras estão ok. \n Exceto a(s) câmera(s):", ', '.join(cameras_filial1))
    else:
        print("· Todas as câmeras do FILIAL - 133 estão ok")

    print("CÂMERAS DA FILIAL - s PRODUÇÃO")
    for channel in range(1, 17):
        url_filial_245 = f"s"
        ok, problema = check_camera(url_filial_245, channel)
        if not ok:
            cameras_filial2.append(problema)
    
    if cameras_filial2:
        print("· Todas as câmeras estão ok. \n Exceto a(s) câmera(s)", ', '.join(cameras_filial2))
    else:
        print("· Todas as câmeras FILIAL - 245 estão ok")

def verificar_backup_nakivo(username, password, sender_email):
    print("\n8 VERIFICAR BACKUP NAKIVO")
    imap_url = 's'
    # Conecta ao servidor IMAP
    mail = imaplib.IMAP4_SSL(imap_url)
    mail.login(username, password)
    mail.select('inbox')  # Escolhe a caixa de entrada
    
    # Procura pelo último e-mail de um remetente específico
    type, data = mail.search(None, f'(FROM "{sender_email}")')
    email_ids = data[0].split()
    
    # Pega o ID do último e-mail (o maior número é o mais recente)
    latest_email_id = email_ids[-1]
    type, data = mail.fetch(latest_email_id, '(RFC822)')
    
    for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject = decode_header(msg['subject'])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            
            print('Subject:', subject)
            
            # Extrai e processa o corpo do e-mail
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type in ["text/plain", "text/html"]:
                        body = part.get_payload(decode=True).decode()
                        extract_info_from_body(body)
            else:
                body = msg.get_payload(decode=True).decode()
                extract_info_from_body(body)
    
    mail.logout()

def extract_info_from_body(body):
    info_sections = {
        "Summary": [],
        "Backups": [],
        "Target Storage": [],
        "Alarms & Notifications": []
    }

    current_section = None

    for line in body.splitlines():
        line = line.strip()
        if line in info_sections:
            current_section = line
        elif current_section:
            info_sections[current_section].append(line)

    for section, lines in info_sections.items():
        print(f"{section}:")
        for line in lines:
            print(line)
        print()

def main():
    hoje = pd.to_datetime('today').normalize()
    nome_arquivo = hoje.strftime("checagem_diaria_%d-%m-%Y_.txt")
    original_stdout = sys.stdout
    sys.stdout = open(nome_arquivo, 'w', encoding='utf-8')
    print("\nChecagem em andamento...\n")
    verificar_backup_webglass('s')
    verificar_site('s')
    verificar_backup_fortes('s')
    verificar_vms()
    verificar_controladoras('s', 'a')
    checar_cameras()
    verificar_backup_nakivo('s', 'a', 'v')
    sys.stdout.close()
    sys.stdout = original_stdout
    root.after(0, finish_process)

# Login
login_window = tk.Tk()
login_window.title("Login")
login_window.geometry("400x200")
label_username = Label(login_window, text="Usuário:")
label_username.pack()
entry_username = Entry(login_window)
entry_username.pack()
label_password = Label(login_window, text="Senha:")
label_password.pack()
entry_password = Entry(login_window, show="*")
entry_password.pack()
button_login = Button(login_window, text="Entrar", command=login)
button_login.pack()
login_window.mainloop()

if __name__ == '__main__':
    login_window.mainloop()
