import psutil
import time
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

# Création d'un fichier Excel et initialisation de la feuille avec mise en forme
def create_excel_file():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Utilisation des ressources"
    
    # Ajout des en-têtes de colonnes
    ws["A1"] = "Date"
    ws["B1"] = "Utilisation de la RAM (Go)"
    ws["C1"] = "Utilisation du CPU (%)"
    ws["D1"] = "Utilisation du réseau (Mo/s)"
    ws["E1"] = "Nombre de connexions HTTP/HTTPS"
    
    # Mise en forme des en-têtes
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = openpyxl.styles.PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.border = Border(bottom=Side(style="thin"))
    
    # Ajustement de la largeur des colonnes
    for col in range(1, 6):
        column_letter = get_column_letter(col)
        ws.column_dimensions[column_letter].width = 20
    
    # Sauvegarde initiale du fichier Excel
    wb.save("Rapport_systeme.xlsx")

# Fonction pour obtenir l'utilisation de la RAM en Go
def get_ram_usage_in_gb():
    ram = psutil.virtual_memory()
    ram_in_gb = ram.used / (1024 ** 3)  # Conversion de l'utilisation de la RAM en Go
    return round(ram_in_gb, 2)

# Fonction pour obtenir l'utilisation du CPU en %
def get_cpu_usage_in_percent():
    return psutil.cpu_percent(interval=1)

# Fonction pour obtenir l'utilisation du réseau en Mo/s
def get_network_usage_in_mbps():
    net_io = psutil.net_io_counters()
    # Obtenir l'usage actuel en octets
    bytes_sent = net_io.bytes_sent
    bytes_recv = net_io.bytes_recv
    
    time.sleep(1)
    
    net_io = psutil.net_io_counters()
    bytes_sent_new = net_io.bytes_sent
    bytes_recv_new = net_io.bytes_recv
    
    # Calculer l'utilisation du réseau en Mo/s (mégaoctets par seconde)
    sent_per_sec = (bytes_sent_new - bytes_sent) / (1024 * 1024)
    recv_per_sec = (bytes_recv_new - bytes_recv) / (1024 * 1024)
    
    return round(sent_per_sec + recv_per_sec, 2)

# Fonction pour obtenir le nombre de connexions HTTP/HTTPS
def get_http_connections_count():
    http_requests = 0
    for conn in psutil.net_connections(kind='inet'):
        # Vérifier les connexions sur les ports 80 (HTTP) ou 443 (HTTPS)
        if conn.status == 'ESTABLISHED' and (conn.laddr.port == 80 or conn.laddr.port == 443):
            http_requests += 1
    return http_requests

# Fonction pour ajouter des graphiques individuels pour chaque donnée
def add_individual_charts(ws):
    # Création des graphiques
    chart_positions = ["G2", "G20", "G38", "G56"]  # Positions des graphiques dans la feuille
    titles = ["Utilisation de la RAM (Go)", "Utilisation du CPU (%)", "Utilisation du réseau (Mo/s)", "Nombre de connexions HTTP/HTTPS"]
    
    for i, col_idx in enumerate(range(2, 6), start=0):  # Les colonnes B à E
        chart = LineChart()
        chart.title = titles[i]
        chart.style = 13
        chart.y_axis.title = titles[i]
        chart.x_axis.title = "Temps"
        chart.height = 10  # Taille du graphique
        chart.width = 20
        
        # Définir les plages de données
        data = Reference(ws, min_col=col_idx, min_row=1, max_row=ws.max_row)
        categories = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
        
        # Ajouter les données et catégories au graphique
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        
        # Ajouter le graphique à la feuille
        ws.add_chart(chart, chart_positions[i])

# Fonction pour appliquer les couleurs sur les valeurs au-dessus et en-dessous de la moyenne
def apply_colors(ws, col_idx):
    values = []
    max_row = ws.max_row

    # Récupérer les valeurs de la colonne
    for row in range(2, max_row + 1):
        value = ws.cell(row=row, column=col_idx).value
        values.append(value)
    
    # Calcul de la moyenne
    average_value = sum(values) / len(values)
    
    # Appliquer les couleurs en fonction de la moyenne
    for row in range(2, max_row + 1):
        cell = ws.cell(row=row, column=col_idx)
        value = cell.value
        
        if value > average_value:
            cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # Rouge clair
        elif value < average_value:
            cell.fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")  # Vert clair

# Fonction pour ajouter les données dans le fichier Excel
def log_system_usage():
    if os.path.exists("Rapport_systeme.xlsx"):
        wb = openpyxl.load_workbook("Rapport_systeme.xlsx")
        ws = wb.active
    else:
        # Si le fichier n'existe pas, le créer
        wb = openpyxl.Workbook()
        ws = wb.active
        create_excel_file()
    
    # Obtenir l'utilisation des différentes ressources
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ram_usage_in_gb = get_ram_usage_in_gb()
    cpu_usage_percent = get_cpu_usage_in_percent()
    network_usage_mbps = get_network_usage_in_mbps()
    http_requests_count = get_http_connections_count()
    
    # Ajouter les données dans la prochaine ligne vide
    new_row = ws.max_row + 1
    ws[f"A{new_row}"] = timestamp
    ws[f"B{new_row}"] = ram_usage_in_gb
    ws[f"C{new_row}"] = cpu_usage_percent
    ws[f"D{new_row}"] = network_usage_mbps
    ws[f"E{new_row}"] = http_requests_count
    
    # Mise en forme des données
    for col in range(1, 6):
        cell = ws.cell(row=new_row, column=col)
        cell.font = Font(name='Calibri', size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    # Appliquer les couleurs sur les valeurs au-dessus et en-dessous de la moyenne
    for col_idx in range(2, 6):
        apply_colors(ws, col_idx)
    
    # Ajouter ou mettre à jour les graphiques individuels
    add_individual_charts(ws)
    
    # Sauvegarder le fichier Excel
    wb.save("Rapport_systeme.xlsx")
    print(f"{timestamp} - RAM: {ram_usage_in_gb} Go, CPU: {cpu_usage_percent}%, Network: {network_usage_mbps} Mo/s, HTTP/HTTPS Requests: {http_requests_count}")

# Vérifier si le fichier Excel existe, sinon le créer
if not os.path.exists("Rapport_systeme.xlsx"):
    create_excel_file()

# Boucle pour exécuter le script toutes les 5 minutes (300 secondes)
while True:
    log_system_usage()
    time.sleep(10)
