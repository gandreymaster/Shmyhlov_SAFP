import pandas as pd
import numpy as np


def create_year_dict(row):
    keys = [
        'Turnover', 'Profit Net', 'Liailities', 'Fixed assets',
        'Circulant Assets', 'Capitals and reserves', 'The average number of employees'
    ]
    year_dict = {}

    try:
        year_value = str(row['Unnamed: 3']).replace('\xa0', '').strip()
        year_dict['Year'] = float(year_value) if year_value != 'nan' else np.nan
    except (ValueError, KeyError) as e:
        year_dict['Year'] = np.nan
        print(f"Помилка при обробці року: {e}")

    for key in keys:
        try:
            value = str(row[key]).replace('\xa0', '').strip()
            year_dict[key] = float(value) if value != 'nan' else np.nan
        except (ValueError, KeyError) as e:
            year_dict[key] = np.nan
            print(f"Помилка при обробці {key} у рядку: {e}")

    return year_dict


def get_activity_list(file_name):
    try:
        xl_file = pd.ExcelFile(file_name)
        df_companies_activity = xl_file.parse('Form Responses 1')
        activity_list = [str(activity).strip() for activity in
                         df_companies_activity['What is your field of activity?'].tolist()]

        activity_mapping = {
            'Activitati de secretariat': {"Services": {"Professional Services": "Secretarial activities"}},
            'IT': {"IT and Technology": "IT"},
            'Domeniul medical': {"Medical and Healthcare": "Medical field"},
            'Servicii': {"Services": {"General Services": "Services"}},
            'Textile': {"Trade": {"Specialized Trade": "Textiles"}},
            'Constructii': {"Construction": {"General Construction": "Construction"}},
            'Produse industriale': {"Trade": {"Specialized Trade": "Industrial products"}},
            'Transport și logistică': {"Transport and Logistics": {"Specialized Transport": "Transport and logistics"}},
            'Comert cu amanuntul al articolelor de fierarie, al articolelor din sticla si a celor pentru vopsit, in magazine specializate (CAEN 4752)':
                {"Trade": {
                    "Retail Trade": "Retail trade of hardware, glassware, and painting supplies in specialized stores (CAEN 4752)"}},
            'Electronice': {"Trade": {"Specialized Trade": "Electronics"}},
            'Cosmetica': {"Services": {"Beauty and Personal Care": "Cosmetics"}},
            'Transport express': {"Transport and Logistics": {"Freight Transport": "Express transport"}},
            'constructii': {"Construction": {"General Construction": "Construction"}},
            'Cosmetice naturale': {"Services": {"Beauty and Personal Care": "Natural cosmetics"}},
            'mobilier': {"Trade": {"Specialized Trade": "Furniture"}},
            'Comert': {"Trade": {"General Trade": "Trade"}},
            'Transport': {"Transport and Logistics": {"General Transport": "Transport"}},
            'transporturi': {"Transport and Logistics": {"General Transport": "Transportation"}},
            'Imobiliare': {"Real Estate": "Real estate"},
            'Alimente și băuturi': {"Trade": {"Specialized Trade": "Food and beverages"}},
            'Transporturi Rutiere de Marfuri': {
                "Transport and Logistics": {"Freight Transport": "Road freight transport"}},
            'real estate - intermedieri': {"Real Estate": "Real estate - intermediation"},
            'Lucrari de terasamente': {"Construction": {"Specialized Construction": "Earthworks"}},
            'TRANSPORT': {"Transport and Logistics": {"General Transport": "Transport"}},
            'Transport marfa': {"Transport and Logistics": {"Freight Transport": "Freight transport"}},
            'servicii': {"Services": {"General Services": "Services"}},
            'Transport rutier de mârfuri.': {
                "Transport and Logistics": {"Freight Transport": "Road freight transport"}},
            'constructii, amenajari interioare, exterioare, vanzari diferite materiale de constructii si amenajari interioare si exterioare':
                {"Construction": {
                    "Specialized Construction": "Construction, interior and exterior design, sales of various construction and interior/exterior design materials"}},
            'SERVICII TRADUCERI': {"Services": {"Professional Services": "Translation services"}},
            'transport': {"Transport and Logistics": {"General Transport": "Transport"}},
            'Transporturi rutiere': {"Transport and Logistics": {"General Transport": "Road transport"}},
            'Voluntariat': {"Volunteering": "Volunteering"},
            'Medical': {"Medical and Healthcare": "Medical field"},
            'Transport Marfuri Generale': {
                "Transport and Logistics": {"Freight Transport": "General freight transport"}},
            'Transport general': {"Transport and Logistics": {"General Transport": "General transport"}},
            'TRANSPORT RUTIER DE MARFURI': {"Transport and Logistics": {"Freight Transport": "Road freight transport"}},
            'Transport rutier': {"Transport and Logistics": {"General Transport": "Road transport"}},
            'Domeniu de frumusete si ingrijire corporala': {
                "Services": {"Beauty and Personal Care": "Beauty and body care field"}},
            'Transporturi': {"Transport and Logistics": {"General Transport": "Transportation"}},
            'Construcții': {"Construction": {"General Construction": "Construction"}},
            'Transport marfuri': {"Transport and Logistics": {"Freight Transport": "Freight transport"}},
            'Horticultura peisagistica': {"Construction": {"Specialized Construction": "Landscape horticulture"}},
            'Asistență medicală': {"Medical and Healthcare": "Medical assistance"},
            'TRANSPORT MARFURI GENERALE': {
                "Transport and Logistics": {"Freight Transport": "General freight transport"}},
            'Pharma': {"Medical and Healthcare": "Pharma"},
            'vanzare- produse de asigurare': {"Insurance": "Sales - insurance products"},
            'Comert cu amanuntul': {"Trade": {"Retail Trade": "Retail trade"}},
            'Trasport': {"Transport and Logistics": {"General Transport": "Transport"}},
            'Vanzarea cu amanuntul': {"Trade": {"Retail Trade": "Retail sales"}},
            'Produse de ingrijire personala': {"Trade": {"Specialized Trade": "Personal care products"}},
            'Turism': {"Services": {"Tourism": "Tourism"}},
            'Transport marfă': {"Transport and Logistics": {"Freight Transport": "Freight transport"}},
            'Servicii financiare': {"Services": {"Professional Services": "Financial services"}},
            'Transport mărfuri': {"Transport and Logistics": {"Freight Transport": "Freight transport"}},
            'COMERT NEALIMENTAR': {"Trade": {"General Trade": "Non-food trade"}},
            'transport marfa': {"Transport and Logistics": {"Freight Transport": "Freight transport"}},
            'Dezvoltare hardware și software': {"IT and Technology": "Hardware and software development"},
            'Automotive': {"Automotive": "Automotive"},
            'Transport marfă': {"Transport and Logistics": {"Freight Transport": "Freight transport"}}
        }
        structured_list = [activity_mapping.get(activity, {"Uncategorized": activity}) for activity in activity_list]

    except Exception as e:
        raise Exception(f"Не вдалося відкрити файл: {e}")

    return structured_list


def read_xlsx(file_name):
    try:
        xl_file = pd.ExcelFile(file_name)
        df_companies = xl_file.parse('Sheet1')
        df_companies_activity = get_activity_list(file_name)
        df_companies_status = xl_file.parse('Form Responses 1')
        df_companies_status = [True if value == 'DA' else False
                  for value in df_companies_status['Do you have external customers/suppliers?']]
    except Exception as e:
        raise Exception(f"Не вдалося відкрити файл: {e}")

    df_companies.columns = df_companies.columns.str.strip()

    dict_companies = []
    temp_dict = {}
    years_arr = []
    ka = 0

    for index, row in df_companies.iterrows():
        ismanufacturing = False
        if row.isna().all():
            if temp_dict:
                temp_dict['Years'] = years_arr
                dict_companies.append(temp_dict)
                temp_dict = {}
                years_arr = []
            continue

        if pd.isna(row.get('Nume of the company')) and not pd.isna(row.get('Unnamed: 3')):
            years_dict = create_year_dict(row)
            years_arr.append(years_dict)
        else:
            if temp_dict:
                temp_dict['Years'] = years_arr
                dict_companies.append(temp_dict)

            try:
                company_activity = df_companies_activity[ka]
                company_status = df_companies_status[ka]
                ka += 1
                temp_dict = {
                    'Name': row['Nume of the company'],
                    'Field of activity': company_activity,
                    'Is manufacturing?': company_status,
                    'Year of establishment': float(str(row['Year of establishment']).replace('\xa0', ''))
                    if not pd.isna(row['Year of establishment']) else np.nan
                }
                years_arr = [create_year_dict(row)]
            except ValueError as e:
                print(f"Помилка при обробці року заснування для {row['Nume of the company']}: {e}")
                temp_dict = {'Name': row['Nume of the company'], 'Year of establishment': np.nan}
                years_arr = [create_year_dict(row)]

    if temp_dict:
        temp_dict['Years'] = years_arr
        dict_companies.append(temp_dict)

    return dict_companies


if __name__ == "__main__":
    file_name = 'Firms_Chestionar.xlsx'
    companies_data = read_xlsx(file_name)
    print(companies_data[0])
