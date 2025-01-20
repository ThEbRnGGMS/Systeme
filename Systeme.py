import psutil
import time
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

# Création d'un fichier Excel et initialisation de la feuille
def create_excel_file(file_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Utilisation des ressources"

    # Ajout des en-têtes de colonnes
    headers = ["Date", "Utilisation de la RAM (Go)", "Utilisation du CPU (%)", "Utilisation du réseau (Mo/s)", "Nombre de connexions HTTP/HTTPS"]
    ws.append(headers)

    # Mise en forme des en-têtes
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.border = Border(bottom=Side(style="thin"))

    # Ajustement de la largeur des colonnes
    for col in range(1, 6):
        column_letter = get_column_letter(col)
        ws.column_dimensions[column_letter].width = 25

    wb.save(file_name)

# Fonction pour obtenir l'utilisation de la RAM en Go
def get_ram_usage_in_gb():
    ram = psutil.virtual_memory()
    return round(ram.used / (1024 ** 3), 2)

# Fonction pour obtenir l'utilisation du CPU en %
def get_cpu_usage_in_percent():
    return psutil.cpu_percent(interval=1)

# Fonction pour obtenir l'utilisation du réseau en Mo/s
def get_network_usage_in_mbps():
    net_io = psutil.net_io_counters()
    bytes_sent = net_io.bytes_sent
    bytes_recv = net_io.bytes_recv

    time.sleep(1)

    net_io = psutil.net_io_counters()
    sent_per_sec = (net_io.bytes_sent - bytes_sent) / (1024 * 1024)
    recv_per_sec = (net_io.bytes_recv - bytes_recv) / (1024 * 1024)

    return round(sent_per_sec + recv_per_sec, 2)

# Fonction pour obtenir le nombre de connexions HTTP/HTTPS
def get_http_connections_count():
    http_count = 0
    for conn in psutil.net_connections(kind='inet'):
        if conn.status == 'ESTABLISHED' and conn.laddr.port in (80, 443):
            http_count += 1
    return http_count

# Fonction pour appliquer des couleurs conditionnelles
def apply_colors(ws, col_idx):
    values = [ws.cell(row=row, column=col_idx).value for row in range(2, ws.max_row + 1)]
    average_value = sum(values) / len(values) if values else 0

    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_idx)
        if cell.value > average_value:
            cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        elif cell.value < average_value:
            cell.fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

# Fonction pour ajouter des graphiques individuels
def add_individual_charts(ws):
    chart_positions = ["G2", "G20", "G38", "G56"]
    titles = ["RAM (Go)", "CPU (%)", "Réseau (Mo/s)", "Connexions HTTP/HTTPS"]

    for i, col_idx in enumerate(range(2, 6), start=0):
        chart = LineChart()
        chart.title = titles[i]
        chart.y_axis.title = titles[i]
        chart.x_axis.title = "Temps"

        data = Reference(ws, min_col=col_idx, min_row=1, max_row=ws.max_row)
        categories = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)

        ws.add_chart(chart, chart_positions[i])

# Fonction pour sauvegarder les données supprimées dans un fichier avec la même mise en forme
def save_deleted_data(deleted_data):
    file_name = "OLD_DONNEE.xlsx"
    if not os.path.exists(file_name):
        create_excel_file(file_name)

    wb = openpyxl.load_workbook(file_name)
    ws = wb.active

    for row in deleted_data:
        ws.append(row)

    for col_idx in range(2, 6):
        apply_colors(ws, col_idx)

    add_individual_charts(ws)
    wb.save(file_name)

# Fonction principale pour enregistrer les données
def log_system_usage(num):
    main_file = "Rapport_systeme.xlsx"
    if not os.path.exists(main_file):
        create_excel_file(main_file)

    wb = openpyxl.load_workbook(main_file)
    ws = wb.active

    max_data_rows = 40
    deleted_data = []

    while ws.max_row > max_data_rows:
        deleted_row = [ws.cell(row=2, column=col).value for col in range(1, 6)]
        deleted_data.append(deleted_row)
        ws.delete_rows(2)

    save_deleted_data(deleted_data)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ram_usage = get_ram_usage_in_gb()
    cpu_usage = get_cpu_usage_in_percent()
    network_usage = get_network_usage_in_mbps()
    http_count = get_http_connections_count()

    ws.append([timestamp, ram_usage, cpu_usage, network_usage, http_count])

    for col_idx in range(2, 6):
        apply_colors(ws, col_idx)

    add_individual_charts(ws)
    wb.save(main_file)
    
    print(f"{num} - {timestamp} - RAM: {ram_usage} Go, CPU: {cpu_usage}%, Réseau: {network_usage} Mo/s, Connexions: {http_count}")
    return num + 1  # Incrémente num et le retourne pour le tour suivant

# Exécution périodique
if __name__ == "__main__":
    num = 1  # Initialisation de num avant la boucle principale
    while True:
        num = log_system_usage(num)  # Passe num à chaque tour et incrémente-le
        time.sleep(1)
