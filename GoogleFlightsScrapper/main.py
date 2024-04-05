from google_flight_analysis.scrape import Scrape, ScrapeObjects
import pandas as pd
from datetime import datetime
from openpyxl import Workbook

class FlightData:
    def __init__(self, direction, data):
        self.direction = direction
        self.departure_datetime = data.iloc[0]
        self.arrival_datetime = data.iloc[1]
        self.origin = data.iloc[2]
        self.destination = data.iloc[3]
        self.airlines = data.iloc[4]
        self.travel_time = data.iloc[5]
        self.price = data.iloc[6]
        self.num_stops = data.iloc[7]
        self.layover = data.iloc[8]
        self.access_date = data.iloc[9]
        self.co2_emission = data.iloc[10]
        self.emission_diff = data.iloc[11]

    def print_flight_data_obj(self):
        print("Direction:", self.direction)
        print("Departure Datetime:", self.departure_datetime)
        print("Arrival Datetime:", self.arrival_datetime)
        print("Origin:", self.origin)
        print("Destination:", self.destination)
        print("Airline(s):", self.airlines)
        print("Travel Time:", self.travel_time)
        print("Price ($):", self.price)
        print("Number of Stops:", self.num_stops)
        print("Layover:", self.layover)
        print("Access Date:", self.access_date)
        print("CO2 Emission (kg):", self.co2_emission)
        print("Emission Diff (%):", self.emission_diff)

    def to_list(self):
        return [self.direction, self.departure_datetime, self.arrival_datetime, self.origin, self.destination,
                self.airlines, self.travel_time, self.price, self.num_stops, self.layover,
                self.access_date, self.co2_emission, self.emission_diff]
def generate_filename(prefix='flight_data', extension='xlsx'):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{prefix}_{timestamp}.{extension}"
    return filename
def write_to_excel(lowest_price_up, lowest_price_down, filename):
    # Create a new Workbook
    wb = Workbook()
    # Select the active worksheet
    ws = wb.active

    # Write column headers
    headers = ['Direction', 'Departure datetime', 'Arrival datetime', 'Origin', 'Destination',
               'Airline(s)', 'Travel Time', 'Price ($)', 'Num Stops', 'Layover',
               'Access Date', 'CO2 Emission (kg)', 'Emission Diff (%)']
    ws.append(headers)

    # Write data for lowest price up
    ws.append(lowest_price_up.to_list())

    # Write data for lowest price down
    ws.append(lowest_price_down.to_list())

    # Save the workbook
    wb.save(filename)
def print_flight_data(flight_data):
    # Set option to display all columns
    pd.set_option('display.max_columns', None)
    # Define the new column names and order
    new_column_names = {
        'Departure datetime': 'Departure Date',
        'Arrival datetime': 'Arrival Date',
        'Origin': 'From',
        'Destination': 'To',
        'Airline(s)': 'Airline',
        'Travel Time': 'Duration',
        'Price ($)': 'Price',
        'Num Stops': 'Number of Stops',
        'Layover': 'Layover Time',
        'Access Date': 'Accessed Date',
        'CO2 Emission (kg)': 'CO2 Emission',
        'Emission Diff (%)': 'Emission Difference'
    }

    # Define the desired column order
    column_order = ['Departure Date', 'Arrival Date', 'From', 'To', 'Airline',
                    'Duration', 'Price', 'Number of Stops', 'Layover Time',
                    'Accessed Date', 'CO2 Emission', 'Emission Difference']

    # Rename columns and reorder them
    flight_data_renamed = flight_data.rename(columns=new_column_names)[column_order]

    # Output
    print("-------Output-----------------------------------------------------------------------------")
    print(flight_data_renamed)
    print("------------------------------------------------------------------------------------------")
def print_lowest_price_table(lowest_price_up, lowest_price_down):
    # Concatenate the 'Up' and 'Down' DataFrames
    lowest_price_table_up = FlightData('up', lowest_price_up)
    lowest_price_table_down = FlightData('down', lowest_price_down)

    # Print the tabular format
    print("Tabular Format:")
    # print(combined_table)
    lowest_price_table_up.print_flight_data_obj()
    lowest_price_table_down.print_flight_data_obj()
    f_name = generate_filename('lowest_prices', 'xlsx')
    write_to_excel(lowest_price_table_up, lowest_price_table_down, f_name)
def find_lowest_price_rows(dataframe):
    # Check if the dataframe is empty
    if dataframe.empty:
        return None, None

    # Split the DataFrame into up and down flights
    up_flights = dataframe[dataframe['Direction'] == 'Up']
    down_flights = dataframe[dataframe['Direction'] == 'Down']

    # Find the row with the lowest price for up and down flights
    lowest_price_up = up_flights.loc[up_flights['Price ($)'].idxmin()]
    lowest_price_down = down_flights.loc[down_flights['Price ($)'].idxmin()]

    return lowest_price_up, lowest_price_down
def save_flight_data_to_excel(flight_data, filename):
    # Define the new column names and order
    new_column_names = {
        'Departure datetime': 'Departure Date',
        'Arrival datetime': 'Arrival Date',
        'Origin': 'From',
        'Destination': 'To',
        'Airline(s)': 'Airline',
        'Travel Time': 'Duration',
        'Price ($)': 'Price',
        'Num Stops': 'Number of Stops',
        'Layover': 'Layover Time',
        'Access Date': 'Accessed Date',
        'CO2 Emission (kg)': 'CO2 Emission',
        'Emission Diff (%)': 'Emission Difference'
    }

    # Define the desired column order
    column_order = ['Departure Date', 'Arrival Date', 'From', 'To', 'Airline',
                    'Duration', 'Price', 'Number of Stops', 'Layover Time',
                    'Accessed Date', 'CO2 Emission', 'Emission Difference']

    # Rename columns and reorder them
    flight_data_renamed = flight_data.rename(columns=new_column_names)[column_order]

    # Save DataFrame to Excel
    flight_data_renamed.to_excel(filename, index=False)
def add_direction_column(dataframe):
    # Check if the DataFrame is empty
    if dataframe.empty:
        return dataframe

    # Initialize an empty list to store the directions
    directions = []

    # Iterate over the DataFrame rows
    for index, row in dataframe.iterrows():
        # Check if the destination airport matches the first origin airport
        if row['Destination'] == dataframe.iloc[0]['Origin']:
            directions.append('Down')
        else:
            directions.append('Up')

    # Add the directions list as a new column to the DataFrame
    dataframe['Direction'] = directions

    return dataframe
def scrape_flight_data(origin, destination, start_date, end_date):
    # Obtain our scrape object, represents our query
    result = Scrape(origin, destination, start_date, end_date)

    # Run selenium through ChromeDriver, modifies results in-place
    ScrapeObjects(result)

    # Return the queried representation of result
    return result.data
def main():
    # Input parameters
    destination = 'BBI'  # Origin
    origin = 'DFW'  # Destination

    start_date = '2024-07-20'
    end_date = '2024-08-20'

    # Scrape flight data
    flight_data = scrape_flight_data(origin, destination, start_date, end_date)

    # Add Direction Column
    updated_flight_data = add_direction_column(flight_data)
    f_name = generate_filename()
    save_flight_data_to_excel(flight_data, f_name)
    print("Data Saved to Excel: \n", f_name)  # Output: flight_data_20220406_115923.xlsx

    # # Lowest Price Up and Down:
    lowest_price_up, lowest_price_down = find_lowest_price_rows(flight_data)

    # Example usage:
    print_lowest_price_table(lowest_price_up, lowest_price_down)

if __name__ == "__main__":
    main()
