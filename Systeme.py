import psutil
import time
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
import socket
import re

# Fonction pour créer un fichier Excel avec des en-têtes bien formatés
def create_excel_file(file_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Utilisation des Ressources"

    headers = ["Date", "RAM Utilisé (GB)", "CPU Utilisé (%)", "Réseau (MB/s)", "Connexions HTTP/HTTPS"]
    ws.append(headers)

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.border = Border(bottom=Side(style="thin"))

    for col in range(1, 6):
        column_letter = get_column_letter(col)
        ws.column_dimensions[column_letter].width = 25

    wb.save(file_name)

# Fonction pour obtenir la RAM utilisée en GB
def get_ram_usage_in_gb():
    ram = psutil.virtual_memory()
    return round(ram.used / (1024 ** 3), 2)

# Fonction pour obtenir l'utilisation du CPU en %
def get_cpu_usage_in_percent():
    return psutil.cpu_percent(interval=1)

# Fonction pour mesurer l'activité réseau en MB/s
def get_network_usage_in_mbps():
    net_io = psutil.net_io_counters()
    bytes_sent = net_io.bytes_sent
    bytes_recv = net_io.bytes_recv

    time.sleep(1)

    net_io = psutil.net_io_counters()
    sent_per_sec = (net_io.bytes_sent - bytes_sent) / (1024 * 1024)
    recv_per_sec = (net_io.bytes_recv - bytes_recv) / (1024 * 1024)

    return round(sent_per_sec + recv_per_sec, 2)

# Fonction pour compter les connexions HTTP/HTTPS
def get_http_connections_count():
    http_count = 0
    try:
        for conn in psutil.net_connections(kind='inet'):
            if conn.status == 'ESTABLISHED' and conn.raddr and conn.raddr.port in (80, 443):
                http_count += 1
    except psutil.AccessDenied:
        print("Autorisation refusée pour accéder aux connexions réseau. Essayez d'exécuter en tant qu'administrateur.")
    return http_count

# Fonction pour effectuer une recherche DNS inversée et obtenir le domaine
def get_domain_from_ip(ip):
    try:
        domain = socket.gethostbyaddr(ip)[0]
        if re.match(r"^(?:\d{1,3}\.){3}\d{1,3}$", domain) or "ip-" in domain:
            return None
        return domain if "." in domain else None
    except (socket.herror, socket.gaierror):
        return None

# Fonction pour obtenir les données des domaines
def get_domain_usage_data():
    domain_data = {}
    try:
        for conn in psutil.net_connections(kind='inet'):
            if conn.status == 'ESTABLISHED' and conn.raddr and conn.raddr.port in (80, 443):
                ip = conn.raddr.ip
                domain = get_domain_from_ip(ip)
                if domain:
                    domain_data[domain] = domain_data.get(domain, 0) + 1
    except psutil.AccessDenied:
        print("Autorisation refusée pour accéder aux connexions réseau.")
    return domain_data

# Fonction pour sauvegarder les données des domaines
def save_domain_usage_data(domain_data, file_name="Rapport_Domaines.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Utilisation des Domaines"

    headers = ["Nom de Domaine", "Nombre de Requêtes"]
    ws.append(headers)

    for domain, count in domain_data.items():
        ws.append([domain, count])

    for col in range(1, 3):
        column_letter = get_column_letter(col)
        ws.column_dimensions[column_letter].width = 30

    wb.save(file_name)

# Fonction principale pour enregistrer les données dans un fichier Excel
def log_system_usage():
    file_name = "Rapport_Systeme.xlsx"
    if not os.path.exists(file_name):
        create_excel_file(file_name)

    wb = openpyxl.load_workbook(file_name)
    ws = wb.active

    # Limite de lignes pour éviter un fichier trop volumineux
    max_data_rows = 40
    while ws.max_row > max_data_rows:
        ws.delete_rows(2)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ram_usage = get_ram_usage_in_gb()
    cpu_usage = get_cpu_usage_in_percent()
    network_usage = get_network_usage_in_mbps()
    http_count = get_http_connections_count()

    domain_data = get_domain_usage_data()

    ws.append([timestamp, ram_usage, cpu_usage, network_usage, http_count])

    for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(file_name)
    save_domain_usage_data(domain_data)
    print(f"{timestamp} - RAM: {ram_usage} GB, CPU: {cpu_usage}%, Réseau: {network_usage} MB/s, Connexions: {http_count}")

# Exécution périodique
if __name__ == "__main__":
    while True:
        log_system_usage()
        time.sleep(1)  # Délai entre les enregistrements
