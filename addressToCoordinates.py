import openpyxl
import googlemaps
import time
import os
import logging
from datetime import datetime

def setup_logging():
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"geocoding_log_{timestamp}.txt")

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger()

def validate_api_key(api_key):
    gmaps = googlemaps.Client(key=api_key)
    try:
        # Attempt a simple geocode to validate the key
        result = gmaps.geocode("1600 Amphitheatre Parkway, Mountain View, CA")
        if result:
            logger.info(f"API Key is valid. Sample geocode successful.")
            return True
        else:
            logger.error("API Key validation failed: Geocode request returned no results.")
            return False
    except googlemaps.exceptions.ApiError as e:
        logger.error(f"API Key validation failed: {str(e)}")
        if "REQUEST_DENIED" in str(e):
            logger.error("Please check your API key and ensure it's correctly set up in the Google Cloud Console.")
            logger.error("1. Visit https://console.cloud.google.com/")
            logger.error("2. Select your project")
            logger.error("3. Go to 'APIs & Services' > 'Credentials'")
            logger.error("4. Check if the API key is correctly created and has the necessary permissions")
            logger.error("5. Ensure the Geocoding API is enabled for your project")
            logger.error("6. If you've recently created the key, wait a few minutes for it to activate")
        return False
    except Exception as e:
        logger.error(f"An unexpected error occurred while validating the API key: {str(e)}")
        return False

def geocode_address(gmaps_client, address, max_retries=1):
    retries = 0
    while retries < max_retries:
        try:
            result = gmaps_client.geocode(address)
            if result:
                location = result[0]['geometry']['location']
                return (location['lat'], location['lng'])
            else:
                logger.warning(f"No results found for address: {address}")
                return None
        except googlemaps.exceptions.ApiError as e:
            logger.error(f"API Error for address {address}: {str(e)}")
            return None
        except Exception as e:
            retries += 1
            logger.warning(f"Error geocoding address: {address}. Retry attempt {retries} of {max_retries}. Error: {str(e)}")
            if retries < max_retries:
                time.sleep(0.1)
            else:
                logger.error(f"Failed to geocode address after {max_retries} attempts: {address}")
                return None

    return None

def main():
    global logger
    logger = setup_logging()

    # Initialize the Google Maps client
    api_key = 'REDACTED'  # Replace with your actual API key

    logger.info(f"Attempting to validate API key: {api_key[:5]}...{api_key[-5:]}")
    if not validate_api_key(api_key):
        logger.error("Exiting due to invalid API key.")
        return

    gmaps = googlemaps.Client(key=api_key)

    # Specify file paths
    base_path = "/Users/buford/dev/aspen-dental-geoencoding"
    input_filename = os.path.join(base_path, "Aspen_LatLong_API Pull Request_11.3.25.xlsx")
    output_filename = os.path.join(base_path, "Aspen_LatLong_API Pull Request_11.3.25.xlsx")

    # Check if input file exists
    if not os.path.exists(input_filename):
        logger.error(f"Error: Input file not found at {input_filename}")
        return

    # Load the workbook and select the active sheet
    try:
        wb = openpyxl.load_workbook(input_filename)
        sheet = wb.active
    except Exception as e:
        logger.error(f"Error opening input file: {e}")
        return

    total_addresses = 0
    successful_geocodes = 0
    failed_geocodes = 0
    skipped_existing = 0

    # Assume columns are: Faccode (A), Address (B), Lat (C), Long (D)
    # Start from row 2 (after headers)
    for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=2):
        faccode = row[0].value
        address = row[1].value
        existing_lat = row[2].value
        existing_long = row[3].value

        # Skip if no address
        if not address:
            continue

        # Skip if already has coordinates
        if existing_lat is not None and existing_long is not None:
            logger.info(f"Row {row_idx}: Skipping (already has coordinates): {address}")
            skipped_existing += 1
            continue

        total_addresses += 1
        address_str = str(address).strip()
        logger.info(f"Row {row_idx}: Geocoding: {address_str}")
        coordinates = geocode_address(gmaps, address_str)

        if coordinates:
            row[2].value = coordinates[0]  # Latitude
            row[3].value = coordinates[1]  # Longitude
            successful_geocodes += 1
            logger.info(f"Row {row_idx}: Success - Lat: {coordinates[0]}, Long: {coordinates[1]}")
        else:
            failed_geocodes += 1
            logger.warning(f"Row {row_idx}: Failed to geocode: {address_str}")

        time.sleep(0.1)

    # Save the updated workbook
    try:
        wb.save(output_filename)
        logger.info(f"\n{'='*60}")
        logger.info(f"Geocoding complete. Results written to '{output_filename}'")
        logger.info(f"Skipped (already had coordinates): {skipped_existing}")
        logger.info(f"Total addresses processed: {total_addresses}")
        logger.info(f"Successful geocodes: {successful_geocodes}")
        logger.info(f"Failed geocodes: {failed_geocodes}")
        logger.info(f"{'='*60}")
    except Exception as e:
        logger.error(f"Error saving output file: {e}")

if __name__ == "__main__":
    main()