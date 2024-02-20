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
        portfolio_manager = soup.find('th', text='Name of the Portfolio Manager').find_next('td').text.strip().upper()
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
        number_of_clients_field: float(clients_count),
        aum_field: float(aum)
    }

    return portfolio_manager_data

def check_portfolio_manager_past_months(df, options, year, month, portfolio_managers):

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

  if month_key >= 4 :
    end = 4
  elif month_key == 3 :
    end = 3
  elif month_key == 2 :
    end = 2
  else :
    return df

  data_list = []
  if portfolio_managers:
    # Create a reverse dictionary
    reverse_dict = {v: k for k, v in options.items()}
    for portfolio_manager in portfolio_managers:
      # Find the key for the desired value using the reverse dictionary
      key = reverse_dict.get(portfolio_manager)
      for i in range(1,end):
        data = get_portfolio_manager_general_information(key, portfolio_manager, str(year), month_dict[month_key -i])
        data_list.append(data)

  if data_list:
    for item in data_list:
      keys = list(item.keys())
      values = list(item.values())
      year = f"{values[1]}"
      month = f"{keys[2].split('-')[0].strip()}"
      if values[-1] != 0:
        df.loc[values[0], (year, month, 'Clients')] = values[2]
        df.loc[values[0], (year, month, 'AUM')] = values[3]

  return df

def add_new_portfolio_manager_data(year,month):

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
    AUM = [float(value) for value in df1[f'{month} - AUM']]

    df = pd.read_excel('SEBI-Portfolio-Manager-General-Information.xlsx', header=[0,1,2], index_col=[0], sheet_name='Sheet1')

    if (year, month, 'AUM') in df.columns :
      before = df[(year, month, 'AUM')]

      df[(year, month, 'Clients')] = Clients
      df[(year, month, 'AUM')] = AUM

      after = df[(year, month, 'AUM')]

      # Find the differences between 'before' and 'after'
      differences = after - before

      # Get the portfolio managers where changes occurred
      changed_managers = differences[differences != 0]

      portfolio_managers = []
      for index, value in changed_managers.items():
          portfolio_managers.append(index)

      if portfolio_managers and len(portfolio_managers)<444:
        df = check_portfolio_manager_past_months(df, options, year, month, list(set(portfolio_managers)))
    else :
      df.insert(2, (year, month, 'Clients'), Clients)
      df.insert(3, (year, month, 'AUM'), AUM)

    return df

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
        # Find the table element
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
        first_row_data = [portfolio_manager_name, 'No. Clients'] + [0] * 7
        second_row_data = [portfolio_manager_name, 'AUM'] + [0] * 7

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
            Domestic_Clients_PF: float(item[2]),
            Domestic_Clients_Corporates: float(item[3]),
            Domestic_Clients_Non_Corporates: float(item[4]),
            Foreign_Clients_Non_Residents: float(item[5]),
            Foreign_Clients_FPI: float(item[6]),
            Foreign_Clients_Others: float(item[7]),
            Total: float(item[8])
        }
        output.append(output_dict)

    return output[0], output[1]

def check_particulars_past_months(df, options, year, month, portfolio_managers):

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

  if month_key >= 4 :
    end = 4
  elif month_key == 3 :
    end = 3
  elif month_key == 2 :
    end = 2
  else :
    return df

  data_list = []
  if portfolio_managers:
    # Create a reverse dictionary
    reverse_dict = {v: k for k, v in options.items()}
    for portfolio_manager in portfolio_managers:
      # Find the key for the desired value using the reverse dictionary
      key = reverse_dict.get(portfolio_manager)
      for i in range(1,end):
        data1, data2 = get_Particulars_data(key, portfolio_manager, str(year), month_dict[month_key -i])
        data_list.append(data1)
        data_list.append(data2)

  if data_list:
    for item in data_list:
      keys = list(item.keys())
      values = list(item.values())
      year = f"{values[2]}"
      month = f"{keys[3].split('-')[0].strip()}"
      if values[-1] != 0:
        if values[1] == 'No. Clients':
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'No. Clients'), (year, month, 'Domestic Clients', 'PF/EPFO')] = values[3]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'No. Clients'), (year, month, 'Domestic Clients', 'Corporates')] = values[4]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'No. Clients'), (year, month, 'Domestic Clients', 'Non-Corporates')] = values[5]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'No. Clients'), (year, month, 'Foreign Clients', 'Non-Residents')] = values[6]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'No. Clients'), (year, month, 'Foreign Clients', 'FPI')] = values[7]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'No. Clients'), (year, month, 'Foreign Clients', 'Others')] = values[8]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'No. Clients'), (year, month, ' ', 'Total')] = values[9]
        else :
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'AUM'), (year, month, 'Domestic Clients', 'PF/EPFO')] = values[3]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'AUM'), (year, month, 'Domestic Clients', 'Corporates')] = values[4]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'AUM'), (year, month, 'Domestic Clients', 'Non-Corporates')] = values[5]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'AUM'), (year, month, 'Foreign Clients', 'Non-Residents')] = values[6]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'AUM'), (year, month, 'Foreign Clients', 'FPI')] = values[7]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'AUM'), (year, month, 'Foreign Clients', 'Others')] = values[8]
          df.loc[(df.index == values[0]) & (df[(' ', ' ', 'Particulars', ' ')] == 'AUM'), (year, month, ' ', 'Total')] = values[9]

  return df

def add_new_particulars_data(year,month):

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
    EPFO = [float(value) for value in df1[f'{month} - Domestic Clients PF/EPFO']]
    Corporates = [float(value) for value in df1[f'{month} - Domestic Clients Corporates']]
    Non_Corporates = [float(value) for value in df1[f'{month} - Domestic Clients Non-Corporates']]
    Non_Residents = [float(value) for value in df1[f'{month} - Foreign Clients Non-Residents']]
    FPI = [float(value) for value in df1[f'{month} - Foreign Clients FPI']]
    Others = [float(value) for value in df1[f'{month} - Foreign Clients Others']]
    Total = [float(value) for value in df1[f'{month} - Total']]

    df = pd.read_excel('SEBI-Portfolio-Manager-Particulars.xlsx', header=[0,1,2,3], index_col=[0], sheet_name='Sheet1')

    if (year, month, ' ', 'Total') in df.columns :
      before = df[(year, month, ' ', 'Total')]

      df[(year, month, 'Domestic Clients', 'PF/EPFO')] = EPFO
      df[(year, month, 'Domestic Clients', 'Corporates')] = Corporates
      df[(year, month, 'Domestic Clients', 'Non-Corporates')] = Non_Corporates
      df[(year, month, 'Foreign Clients', 'Non-Residents')] = Non_Residents
      df[(year, month, 'Foreign Clients', 'FPI')] = FPI
      df[(year, month, 'Foreign Clients', 'Others')] = Others
      df[(year, month, ' ', 'Total')] = Total

      after = df[(year, month, ' ', 'Total')]

      # Find the differences between 'before' and 'after'
      differences = after - before

      # Get the portfolio managers where changes occurred
      changed_managers = differences[differences != 0]

      portfolio_managers = []
      # Print the portfolio managers where changes occurred
      for index, value in changed_managers.items():
          portfolio_managers.append(index)

      if portfolio_managers and len(portfolio_managers)<444:
        df = check_particulars_past_months(df, options, year, month, list(set(portfolio_managers)))
    else :
      df.insert(1, (year, month, 'Domestic Clients', 'PF/EPFO'), EPFO)
      df.insert(2, (year, month, 'Domestic Clients', 'Corporates'), Corporates)
      df.insert(3, (year, month, 'Domestic Clients', 'Non-Corporates'), Non_Corporates)
      df.insert(4, (year, month, 'Foreign Clients', 'Non-Residents'), Non_Residents)
      df.insert(5, (year, month, 'Foreign Clients', 'FPI'), FPI)
      df.insert(6, (year, month, 'Foreign Clients', 'Others'), Others)
      df.insert(7, (year, month, ' ', 'Total'), Total)

    return df

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
        # Find the table element
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
        Equity_Listed: float(first_row_data[2]),
        Equity_Unlisted: float(first_row_data[3]),
        Mutual_Funds: float(first_row_data[-3]),
        Others: float(first_row_data[-2]),
        Total: float(first_row_data[-1])
    }

    return output_dict

def check_investment_past_months(df, options, year, month, portfolio_managers):

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

  if month_key >= 4 :
    end = 4
  elif month_key == 3 :
    end = 3
  elif month_key == 2 :
    end = 2
  else :
    return df

  data_list = []
  if portfolio_managers:
    # Create a reverse dictionary
    reverse_dict = {v: k for k, v in options.items()}
    for portfolio_manager in portfolio_managers:
      # Find the key for the desired value using the reverse dictionary
      key = reverse_dict.get(portfolio_manager)
      for i in range(1,end):
        data = get_Investment_data(key, portfolio_manager, str(year), month_dict[month_key -i])
        data_list.append(data)

  if data_list:
    for item in data_list:
      keys = list(item.keys())
      values = list(item.values())
      year = f"{values[2]}"
      month = f"{keys[3].split('-')[0].strip()}"
      if values[-1] != 0:
        df.loc[values[0], (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Listed')] = values[3]
        df.loc[values[0], (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Unlisted')] = values[4]
        df.loc[values[0], (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Mutual Funds', ' ')] = values[5]
        df.loc[values[0], (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Others', ' ')] = values[6]
        df.loc[values[0], (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Total', ' ')] = values[7]

  return df

def add_new_investment_data(year,month):
  # To load the dictionary object from the file
  with open('options.pkl', 'rb') as file:
      options = pickle.load(file)

  year = str(year)
  portfolio_managers_data = []

  for index, (key, value) in enumerate(options.items(), start=1):
      progress_percentage = (index / len(options)) * 100
      print(f'Processing {index}/{len(options)} ({progress_percentage:.2f}%)')

      portfolio_manager_data = get_Investment_data(key, value, year,str(month))

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

  if (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Total', ' ') in df.columns :
    before = df[(year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Total', ' ')]

    df[(' ', ' ', 'Investment Approach', ' ', ' ')] = investment_approach
    df[(year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Listed')] = equity_listed
    df[(year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Unlisted')] = equity_unlisted
    df[(year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Mutual Funds', ' ')] = mutual_funds
    df[(year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Others', ' ')] = others
    df[(year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Total', ' ')] = total

    after = df[(year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Total', ' ')]

    # Find the differences between 'before' and 'after'
    differences = after - before

    # Get the portfolio managers where changes occurred
    changed_managers = differences[differences != 0]

    portfolio_managers = []
    # Print the portfolio managers where changes occurred
    for index, value in changed_managers.items():
        portfolio_managers.append(index)

    if portfolio_managers and len(portfolio_managers)<444 :
      df = check_investment_past_months(df, options, year, month, portfolio_managers)
  else :
    df[(' ', ' ', 'Investment Approach', ' ', ' ')] = investment_approach
    df.insert(1, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Listed'), equity_listed)
    df.insert(2, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Equity', 'Unlisted'), equity_unlisted)
    df.insert(3, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Mutual Funds', ' '), mutual_funds)
    df.insert(4, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Others', ' '), others)
    df.insert(5, (year, month, '(AUM) as on last day of the month (Amount in INR crores)', 'Total', ' '), total)

  return df

#Examples

# For SEBI-Portfolio-Manager-General-Information.xlsx
# Add Sep Data for the general innformation
# df = add_new_portfolio_manager_data(2023,'Sep')
# df.to_excel('SEBI-Portfolio-Manager-General-Information.xlsx')

# For SEBI-Portfolio-Manager-Particulars.xlsx
# Add Sep Data for the Particulars
# df = add_new_particulars_data(2023,'Sep')
# df.to_excel('SEBI-Portfolio-Manager-Particulars.xlsx')

# For SEBI-Portfolio-Manager-Investment-Approach.xlsx
# Add Sep Data for the Investment
# df = add_new_investment_data(2023,'Sep')
# df.to_excel('SEBI-Portfolio-Manager-Investment-Approach.xlsx')
