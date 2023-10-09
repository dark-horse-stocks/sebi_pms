import pickle
import requests
import pandas as pd
from bs4 import BeautifulSoup

def get_portfolio_manager_general_information(portfolio_manager_value, portfolio_manager_name, year, month):
    month_dict = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }

    # Find the key for the given month name
    month_key = None
    for key, value in month_dict.items():
        if value == month:
            month_key = key
            break
    
    url = f'https://www.sebi.gov.in/sebiweb/other/OtherAction.do?doPmr=yes&pmrId={portfolio_manager_value}&year={int(year)}&month={int(month_key)}'
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    try:
        portfolio_manager = soup.find('th', text='Name of the Portfolio Manager').find_next('td').text.strip()
    except AttributeError:
        portfolio_manager = portfolio_manager_name

    try:
        clients_count = soup.find('th', text='No. of clients as on last day of the month').find_next('td').text.strip()
    except AttributeError:
        clients_count = 0

    try:
        aum = soup.find('th', text='Total Assets under Management (AUM) as on last day of the month (Amount in INR crores)').find_next('td').text.strip()
    except AttributeError:
        aum = 0

    number_of_clients_field = f'{month} - Clients'
    aum_field = f'{month} - AUM'

    portfolio_manager_data = {
        'Portfolio Manager': portfolio_manager,
        "Year": year,
        number_of_clients_field: clients_count,
        aum_field: aum
    }

    return portfolio_manager_data

def add_new_general_information_data(year,month):
    """
    Scrape data for the given month and year and then add them to the SEBI Portfolio Manager General Information.

    Args:
    - year (int): the year to scrape data for
    - month (str): the month abbreviation (Jan, Feb, ....) to scrape data for.

    Returns:
    - None: the function does not return anything, but it updates the excel file with the new month data.
    """

    year = str(year)
    month = str(month)
    # To load the dictionary object from the file
    with open('options.pkl', 'rb') as file:
        options = pickle.load(file)

    portfolio_managers_data = []

    for index, (key, value) in enumerate(options.items(), start=1):
        progress_percentage = (index / len(options)) * 100
        print(f'Processing {index}/{len(options)} ({progress_percentage:.2f}%)')

        portfolio_manager_data = get_portfolio_manager_general_information(key, value, year, month)

        portfolio_managers_data.append(portfolio_manager_data)

    # Convert the list of dictionaries to a pandas DataFrame
    df1 = pd.DataFrame(portfolio_managers_data)

    # Extract relevant columns from the DataFrame
    Clients = [float(value) for value in df1[f'{month} - Clients']]
    aum = [float(value) for value in df1[f'{month} - AUM']]

    df = pd.read_excel('SEBI-Portfolio-Manager-General-Information.xlsx', header=[0,1,2], index_col=[0], sheet_name='Sheet1')

    df.insert(2, (year, month, 'Clients'), Clients)
    df.insert(3, (year, month, 'AUM'), aum)

    df.to_excel('SEBI-Portfolio-Manager-General-Information.xlsx')


def get_Particulars_data(portfolio_manager_value, portfolio_manager_name, year, month):

    month_dict = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }
    # Find the key for the given month name
    month_key = None
    for key, value in month_dict.items():
        if value == month:
            month_key = key
            break
    
    url = f'https://www.sebi.gov.in/sebiweb/other/OtherAction.do?doPmr=yes&pmrId={portfolio_manager_value}&year={int(year)}&month={int(month_key)}'
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    try:
        # Find the table element using its class
        table = soup.find('th', text='Domestic Clients').find_parent('table', class_='table table-striped table-bordered table-hover background statistics-table')

        # Extract data from the first and second rows of the tbody
        tbody_rows = table.find("tbody").find_all("tr")

        # Extract data for the first row
        first_row_data = [portfolio_manager_name, 'No. Clients']
        for cell in tbody_rows[0].find_all(["th", "td"])[1:]:
            first_row_data.append(cell.text.strip())

        # Extract data for the second row
        second_row_data = [portfolio_manager_name, 'AUM']
        for cell in tbody_rows[1].find_all(["th", "td"])[1:]:
            second_row_data.append(cell.text.strip())

    except AttributeError:
        first_row_data = [portfolio_manager_name, 'No. Clients'] + [None] * 7
        second_row_data = [portfolio_manager_name, 'AUM'] + [None] * 7

    data = [first_row_data, second_row_data]
    # Transforming the list to the desired format
    output = []

    Domestic_Clients_PF = f'{month} - Domestic Clients PF/EPFO'
    Domestic_Clients_Corporates = f'{month} - Domestic Clients Corporates'
    Domestic_Clients_Non_Corporates = f'{month} - Domestic Clients Non-Corporates'
    Foreign_Clients_Non_Residents = f'{month} - Foreign Clients Non-Residents'
    Foreign_Clients_FPI = f'{month} - Foreign Clients FPI'
    Foreign_Clients_Others = f'{month} - Foreign Clients Others'
    Total = f'{month} - Total'

    for item in data:
        output_dict = {
            "Portfolio Manager": item[0],
            "Particulars": item[1],
            "Year": year,
            Domestic_Clients_PF: item[2],
            Domestic_Clients_Corporates: item[3],
            Domestic_Clients_Non_Corporates: item[4],
            Foreign_Clients_Non_Residents: item[5],
            Foreign_Clients_FPI: item[6],
            Foreign_Clients_Others: item[7],
            Total: item[8]
        }
        output.append(output_dict)

    return output[0], output[1]

def add_new_particulars_data(year,month):
    """
    Scrape data for the given month and year and then add them to the SEBI Portfolio Manager Particulars.

    Args:
    - year (int): the year to scrape data for
    - month (str): the month abbreviation (Jan, Feb, ....) to scrape data for.

    Returns:
    - None: the function does not return anything, but it updates the excel file with the new month data.
    """

    year = str(year)
    month = str(month)
    # To load the dictionary object from the file
    with open('options.pkl', 'rb') as file:
        options = pickle.load(file)

    portfolio_managers_data = []

    for index, (key, value) in enumerate(options.items(), start=1):
        progress_percentage = (index / len(options)) * 100
        print(f'Processing {index}/{len(options)} ({progress_percentage:.2f}%)')

        result_0, result_1 = get_Particulars_data(key, value, year,month)

        portfolio_managers_data.append(result_0)
        portfolio_managers_data.append(result_1)
        
    # Convert the list of dictionaries to a pandas DataFrame
    df1 = pd.DataFrame(portfolio_managers_data)

    # Extract relevant columns from the DataFrame
    pf_epfo = [float(value) for value in df1[f'{month} - Domestic Clients PF/EPFO']]
    Corporates = [float(value) for value in df1[f'{month} - Domestic Clients Corporates']]
    Non_Corporates = [float(value) for value in df1[f'{month} - Domestic Clients Non-Corporates']]
    Non_Residents = [float(value) for value in df1[f'{month} - Foreign Clients Non-Residents']]
    fpi = [float(value) for value in df1[f'{month} - Foreign Clients FPI']]
    Others = [float(value) for value in df1[f'{month} - Foreign Clients Others']]
    Total = [float(value) for value in df1[f'{month} - Total']]

    df = pd.read_excel('SEBI-Portfolio-Manager-Particulars.xlsx', header=[0,1,2,3], index_col=[0], sheet_name='Sheet1')

    df.insert(1, (year, month, 'Domestic Clients', 'PF/EPFO'), pf_epfo)
    df.insert(2, (year, month, 'Domestic Clients', 'Corporates'), Corporates)
    df.insert(3, (year, month, 'Domestic Clients', 'Non-Corporates'), Non_Corporates)
    df.insert(4, (year, month, 'Foreign Clients', 'Non-Residents'), Non_Residents)
    df.insert(5, (year, month, 'Foreign Clients', 'FPI'), fpi)
    df.insert(6, (year, month, 'Foreign Clients', 'Others'), Others)
    df.insert(7, (year, month, ' ', 'Total'), Total)

    df.to_excel('SEBI-Portfolio-Manager-Particulars.xlsx')

def get_Investment_data(portfolio_manager_value, portfolio_manager_name, year, month):
    month_dict = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
    }

    # Find the key for the given month name
    month_key = None
    for key, value in month_dict.items():
        if value == month:
            month_key = key
            break

    url = f'https://www.sebi.gov.in/sebiweb/other/OtherAction.do?doPmr=yes&pmrId={portfolio_manager_value}&year={int(year)}&month={int(month_key)}'
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    try:
        # Find the table element using its class
        table = soup.find('th', text='Investment Approach').find_parent('table', class_='table table-striped table-bordered table-hover background statistics-table')

        # Extract data from the first and second rows of the tbody
        tbody_rows = table.find("tbody").find("tr")

        # Extract data for the first row
        first_row_data = [portfolio_manager_name]
        for cell in tbody_rows.find_all(["th", "td"]):
            first_row_data.append(cell.text.strip())

    except AttributeError:
        first_row_data = [portfolio_manager_name] + [0] * 6

    Equity_Listed = f'{month} - (AUM) as on last day of the month (Amount in INR crores) Equity Listed'
    Equity_Unlisted = f'{month} - (AUM) as on last day of the month (Amount in INR crores) Equity Unlisted'
    Mutual_Funds = f'{month} - (AUM) as on last day of the month (Amount in INR crores) Mutual Funds'
    Others = f'{month} - (AUM) as on last day of the month (Amount in INR crores) Others'
    Total = f'{month} - (AUM) as on last day of the month (Amount in INR crores) Total'

    output_dict = {
        "Portfolio Manager": first_row_data[0],
        "Investment Approach": first_row_data[1],
        "Year": year,
        Equity_Listed: first_row_data[2],
        Equity_Unlisted: first_row_data[3],
        Mutual_Funds: first_row_data[-3],
        Others: first_row_data[-2],
        Total: first_row_data[-1]
    }

    return output_dict

def add_new_investment_data(year,month):
    """
    Scrape data for the given month and year and then add them to the SEBI Portfolio Manager Investment.

    Args:
    - year (int): the year to scrape data for
    - month (str): the month abbreviation (Jan, Feb, ....) to scrape data for.

    Returns:
    - None: the function does not return anything, but it updates the excel file with the new month data.
    """

    year = str(year)
    month = str(month)
    # To load the dictionary object from the file
    with open('options.pkl', 'rb') as file:
        options = pickle.load(file)

    portfolio_managers_data = []

    for index, (key, value) in enumerate(options.items(), start=1):
        progress_percentage = (index / len(options)) * 100
        print(f'Processing {index}/{len(options)} ({progress_percentage:.2f}%)')

        portfolio_manager_data = get_Investment_data(key, value, year,month)

        portfolio_managers_data.append(portfolio_manager_data)

    # Convert the list of dictionaries to a pandas DataFrame
    df1 = pd.DataFrame(portfolio_managers_data)

    # Extract relevant columns from the DataFrame
    investment_approach = [value for value in df1['Investment Approach']]
    equity_listed = [float(value) for value in df1[f'{month} - (AUM) as on last day of the month (Amount in INR crores) Equity Listed']]
    equity_unlisted = [float(value) for value in df1[f'{month} - (AUM) as on last day of the month (Amount in INR crores) Equity Unlisted']]
    mutual_funds = [float(value) for value in df1[f'{month} - (AUM) as on last day of the month (Amount in INR crores) Mutual Funds']]
    others = [float(value) for value in df1[f'{month} - (AUM) as on last day of the month (Amount in INR crores) Others']]
    total = [float(value) for value in df1[f'{month} - (AUM) as on last day of the month (Amount in INR crores) Total']]

    df = pd.read_excel('SEBI-Portfolio-Manager-Investment-Approach.xlsx', header=[0,1,2,3,4], index_col=[0], sheet_name='Sheet1')

    df[(' ', ' ', 'Investment Approach', ' ', ' ')] = investment_approach
    df.insert(1, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Listed'), equity_listed)
    df.insert(2, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Unlisted'), equity_unlisted)
    df.insert(3, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Mutual Funds', ' '), mutual_funds)
    df.insert(4, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Others', ' '), others)
    df.insert(5, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Total', ' '), total)

    df.to_excel('SEBI-Portfolio-Manager-Investment-Approach.xlsx')


#Examples

# For SEBI-Portfolio-Manager-General-Information.xlsx
# Add Sep Daat for the general innformation
# add_new_general_information_data(2023,'Sep')

# For SEBI-Portfolio-Manager-Particulars.xlsx
# Add Sep Daat for the Particulars
# add_new_particulars_data(2023,'Sep')

# For SEBI-Portfolio-Manager-Investment-Approach.xlsx
# Add Sep Daat for the Investment
# add_new_investment_data(2023,'Sep')