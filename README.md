# Address to Coordinates Converter

This Python script uses the Google Maps Geocoding API to convert a list of addresses stored in an Excel file (`addresses.xlsx`) to their corresponding latitude and longitude coordinates. The resulting coordinates are saved to a new Excel file (`address_coordinates.xlsx`).

## Prerequisites

- Python 3.x
- openpyxl library
- os module
- logging module
- datetime module

## Usage

1. Create a `config.py` file in the same directory as the `googleApiAddressToCoord.py` file, and add your Google Maps API key to it as a variable named `GOOGLE_MAPS_API_KEY`.
2. Ensure the `addresses.xlsx` file is in the same directory as the `googleApiAddressToCoord.py` file.
3. Run the `googleApiAddressToCoord.py` script.
4. The `address_coordinates.xlsx` file will be created in the same directory, containing the original addresses and their corresponding latitude and longitude coordinates.

## Notes

- The script assumes the addresses are stored in the 'Address' column of the `addresses.xlsx` file.
- If an address cannot be geocoded, the latitude and longitude columns will be filled with `None` values.
- This script is provided as-is and without any warranty. Use it at your own risk.

## Contributing

If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.
