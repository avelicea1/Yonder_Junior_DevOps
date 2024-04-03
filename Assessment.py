import requests
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font


def fetch_data(api_url, length=150):
    response = requests.get(f"{api_url}?length={length}")
    data = response.json()
    return data


def list_suspended_licenses(data):
    suspended_licenses = [license for license in data if license['suspendat']]
    return suspended_licenses


def extract_valid_licenses(data):
    today = datetime.now().strftime("%d/%m/%Y")
    valid_licenses = [license for license in data if license['dataDeExpirare'] >= today]
    return valid_licenses


def find_license_count_by_category(data):
    license_counts = {}
    for license in data:
        category = license['categorie']
        if category in license_counts:
            license_counts[category] += 1
        else:
            license_counts[category] = 1
    return license_counts


def generate_excel_report(data, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Nume", "Prenume", "Categorie", "Data de Emitere", "Data de Expirare", "Suspendat"])

    for license in data:
        ws.append([license['id'], license['nume'], license['prenume'], license['categorie'], license['dataDeEmitere'],
                   license['dataDeExpirare'], license['suspendat']])

    bold_font = Font(bold=True)
    for cell in ws["1:1"]:
        cell.font = bold_font

    wb.save(filename)


if __name__ == "__main__":
    api_url = "http://localhost:30000/drivers-licenses/list"
    data = fetch_data(api_url, length=150)

    while True:
        print("\nOperations:")
        print("1. List suspended licenses")
        print("2. Extract valid licenses issued until today's date")
        print("3. Find licenses based on category and their count")
        print("0. Exit")

        operation = input("\nEnter operation ID: ")

        if operation == "1":
            suspended_licenses = list_suspended_licenses(data)
            print("Suspended licenses:")
            for license in suspended_licenses:
                print(license)

        elif operation == "2":
            valid_licenses = extract_valid_licenses(data)
            print("Valid licenses issued until today's date:")
            for license in valid_licenses:
                print(license)

        elif operation == "3":
            license_counts = find_license_count_by_category(data)
            print("License counts by category:")
            for category, count in license_counts.items():
                print(f"{category}: {count}")

        elif operation == "0":
            print("Exiting program.")
            break

        else:
            print("Invalid operation ID. Please enter a valid operation ID.")

    generate_excel_report(data, "licenses_report.xlsx")
