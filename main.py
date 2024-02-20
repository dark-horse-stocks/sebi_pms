import sys
from app import portfolio_manager_data, particulars_data, investment_data

def main():
    # Check if all params are provided as command-line arguments
    if len(sys.argv) != 4:
        print("Usage: python main.py <file> <year> <month>")
        print("file: can take three possible values")
        print("values: portfolio, particulars, investments")
        sys.exit(1)

    # Parse command-line arguments
    file = sys.argv[1]
    year = int(sys.argv[2])
    month = sys.argv[3]

    if file == 'portfolio':
      # Call the portfolio_manager_data function
      portfolio_manager_data(year, month)
    elif file == 'particulars':
      # Call the particulars_data function
      particulars_data(year, month)
    elif file == 'investments':
      # Call the investment_data function
      investment_data(year, month)
       

if __name__ == "__main__":
    main()
