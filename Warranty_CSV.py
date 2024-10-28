import requests
import json
import csv
from datetime import datetime
import re

# Function to get the machine type from the serial number
def get_type_info(serial_number):
    url = f"https://pcsupport.lenovo.com/us/en/api/v4/mse/getproducts?productId={serial_number}"
    response = requests.get(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

    try:
        response_data = response.json()
    except json.JSONDecodeError:
        print("Failed to decode JSON from response.")
        return None  # Return None if the JSON response is invalid

    # Check if the response is a list and retrieve the first element if so
    if isinstance(response_data, list) and len(response_data) > 0:
        type_info = response_data[0].get("Name", "")
    elif isinstance(response_data, dict):
        type_info = response_data.get("Name", "")
    else:
        print("Unexpected response format:", response_data)
        return None  # Return None if the format is not as expected
    # Improved extraction logic for the single type number and type name
    type_number = None
    type_name = None

    # Use regex to find the last occurrence of "Type "
    import re
    type_matches = re.findall(r'Type (\w+)', type_info)
    if type_matches:
        type_number = type_matches[-1]  # Get the last matched type number

    type_name = type_info.split(' - ')[0] if ' - ' in type_info else None

    return {
        "SerialNumber": serial_number,
        "FullType": type_info,
        "TypeNumber": type_number,
        "TypeName": type_name
    }


# Function to get warranty details using the serial number and machine type
def get_warranty_info(serial_number, machine_type):
    url = "https://pcsupport.lenovo.com/us/en/api/v4/upsell/redport/getIbaseInfo"
    payload = json.dumps({
        "serialNumber": serial_number,
        "machineType": machine_type,
        "country": "us",
        "language": "en"
    })
    headers = {
        'accept': 'application/json, text/plain, */*',
        'content-type': 'application/json',
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    response = requests.post(url, headers=headers, data=payload)
    return response.json()


# Function to write warranty details to CSV
def write_warranty_to_csv(csv_writer, warranty, status):
    csv_writer.writerow([
        warranty.get('name', 'N/A'),
        warranty.get('type', 'N/A'),
        warranty.get('description', 'N/A'),
        warranty.get('duration', 'N/A'),
        warranty.get('startDate', 'N/A'),
        warranty.get('endDate', 'N/A'),
        warranty.get('deliveryTypeName', 'N/A'),
        warranty.get('level', 'N/A'),
        status
    ])


# Function to print enhanced warranty and machine information
def print_enhanced_info(warranty_details, csv_writer):
    machine_info = warranty_details['data']['machineInfo']
    base_warranties = warranty_details['data']['baseWarranties']
    upgrade_warranties = warranty_details['data']['upgradeWarranties']

    # Write basic product information to CSV
    csv_writer.writerow([
        machine_info['serial'],
        machine_info['model'],
        machine_info['productName'],
        machine_info['buildDate'],
        machine_info['shipToCountry'],
        machine_info['status'],
        machine_info['brand'],
        machine_info['series'],
        machine_info['productImage'],
        warranty_details['data'].get('warrantyStatus', 'N/A'),
        'Yes' if warranty_details['data'].get('oow', False) else 'No'
    ])

    # Get current date for comparison
    today = datetime.now().date()

    # Process base warranties
    if base_warranties:
        for warranty in base_warranties:
            end_date = datetime.strptime(warranty.get('endDate', '9999-12-31'), '%Y-%m-%d').date()
            status = 'Active' if end_date >= today else 'Expired'
            write_warranty_to_csv(csv_writer, warranty, status)

    # Process upgrade warranties
    if upgrade_warranties:
        for warranty in upgrade_warranties:
            end_date = datetime.strptime(warranty.get('endDate', '9999-12-31'), '%Y-%m-%d').date()
            status = 'Active' if end_date >= today else 'Expired'
            write_warranty_to_csv(csv_writer, warranty, status)

    # Current warranty
    if 'currentWarranty' in warranty_details['data']:
        current_warranty = warranty_details['data']['currentWarranty']
        end_date = datetime.strptime(current_warranty.get('endDate', '9999-12-31'), '%Y-%m-%d').date()
        status = 'Active' if end_date >= today else 'Expired'
        write_warranty_to_csv(csv_writer, current_warranty, status)


# Example usage
serial_number = input("Enter the serial number: ")  # Get serial number from user
type_info = get_type_info(serial_number)

if type_info and type_info["TypeNumber"]:
    machine_type = type_info["TypeNumber"]
    warranty_details = get_warranty_info(serial_number, machine_type)

    # Sanitize the serial number for the filename
    sanitized_serial = re.sub(r'[<>:"/\\|?*]', '_', serial_number)  # Replace invalid characters with underscores
    csv_filename = f'Warranty_info_{sanitized_serial}.csv'  # Create a CSV filename

    # Open CSV file for writing
    with open(csv_filename, mode='w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        # Write header
        csv_writer.writerow([
            "Serial Number", "Model", "Product Name", "Build Date",
            "Ship-To Location", "Status", "Brand", "Series",
            "Product Image", "Warranty Status", "Out of Warranty",
            
        ])
        
        # Print and save enhanced info
        print_enhanced_info(warranty_details, csv_writer)
else:
    print("Unable to retrieve machine type.")
