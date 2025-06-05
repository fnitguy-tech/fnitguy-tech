# This script was created by David Osborne @ https://www.linkedin.com/in/davidrayosborne/ and modified by Zachary Rogers @ https://www.linkedin.com/in/fnitguy/

import os
import csv
from zeep import Client
from zeep.transports import Transport
from requests import Session
from requests.auth import HTTPBasicAuth
import urllib3
import subprocess

import os
import csv
from zeep import Client
from zeep.transports import Transport
from requests import Session
from requests.auth import HTTPBasicAuth
import urllib3
import subprocess

# Suppress SSL certificate warnings (for self-signed CUCM certs)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- Configuration Section ---

# CUCM AXL API URL and credentials
cucm_url = "CUCM IP ADDRESS:8443/axl/"  # Replace with actual CUCM AXL endpoint
username = "CUCM USERNAME"
password = "PASSWORD"

# Path to local WSDL file (defines the SOAP API contract for CUCM AXL)
wsdl_url = r"WSDL FILE PATH\AXLAPI.wsdl"

# Create a requests session for authentication and disabling SSL verification
session = Session()
session.verify = False  # Ignores SSL certificate validation
session.auth = HTTPBasicAuth(username, password)

# Attach session to Zeep transport layer (for SOAP calls)
transport = Transport(session=session)

# Initialize Zeep SOAP client with WSDL and transport
client = Client(wsdl=wsdl_url, transport=transport)

# Paths to CSV files used for input/output
base_path = r"PATH FOR CSV FILES"
user_csv_file = os.path.join(base_path, "user_data.csv")  # Output file for CUCM user data
mappings_csv_file = os.path.join(base_path, "mappings.csv")  # Input file for prefix mappings

# --- Load CSV Mappings (Prefix → CUCM settings) ---

# Initialize dictionaries for prefix-based mappings
device_pool_mapping = {}
route_partition_mapping = {}
calling_search_space_mapping = {}

# Read the mappings file and populate dictionaries
try:
    with open(mappings_csv_file, 'r') as csvfile:
        csv_reader = csv.DictReader(csvfile)
        for row in csv_reader:
            prefix = row['Prefix'].strip()
            device_pool_mapping[prefix] = row['DevicePool'].strip()
            route_partition_mapping[prefix] = row['RoutePartition'].strip()
            calling_search_space_mapping[prefix] = row['CallingSearchSpace'].strip()
    print("Mappings loaded successfully from CSV file.")
except FileNotFoundError:
    print(f"Error: The file {mappings_csv_file} was not found.")
    exit()
except Exception as e:
    print(f"An error occurred while reading the mappings CSV file: {e}")
    exit()

# --- Helper Functions ---

# Return device pool based on 3-digit prefix of phone number
def get_device_pool(phone_number):
    return device_pool_mapping.get(phone_number[:3], 'DefaultDevicePool')

# Return route partition based on prefix
def get_route_partition(phone_number):
    return route_partition_mapping.get(phone_number[:3], 'DefaultRoutePartition')

# Return calling search space based on prefix
def get_calling_search_space(phone_number):
    return calling_search_space_mapping.get(phone_number[:3], 'Default_CSS')

# --- Query CUCM for User Info and Save to CSV ---

try:
    # SQL query to retrieve user and phone data from CUCM
    sql_query = "SELECT userid, keypadEnteredAlternateIdentifier, telephonenumber, firstName, lastName FROM enduser"
    response = client.service.executeSQLQuery(sql=sql_query)

    # Parse returned rows from CUCM
    rows = response['return'].row if response['return'] and hasattr(response['return'], 'row') else []

    # Write extracted user data to CSV for use in later steps
    with open(user_csv_file, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(["UserID", "PhoneNumber", "FirstName", "LastName"])  # CSV header

        for row in rows:
            user_id = row[0].text if row[0] is not None else ''
            if user_id.startswith('Token_User_'):
                continue  # Skip CUCM system or temp users

            # Determine which number to use
            keypad_number = row[1].text if row[1] is not None else ''
            telephone_number = row[2].text if row[2] is not None else ''
            phone_number = keypad_number or telephone_number

            # Get names
            first_name = row[3].text if row[3] is not None else ''
            last_name = row[4].text if row[4] is not None else ''

            # Write user data row
            csv_writer.writerow([user_id, phone_number, first_name, last_name])

    print(f"User data saved to {user_csv_file}")
except Exception as e:
    print(f"Failed to retrieve user data: {e}")

# --- Read User Data from CSV and Add CUCM Phones ---

try:
    with open(user_csv_file, 'r') as csvfile:
        csv_reader = csv.DictReader(csvfile)
        for row in csv_reader:
            user_id = row['UserID']
            phone_number = row['PhoneNumber']
            first_name = row['FirstName']
            last_name = row['LastName']
            full_name = f"{first_name} {last_name}".strip()

            # Look up CUCM configs based on number prefix
            device_pool = get_device_pool(phone_number)
            route_partition = get_route_partition(phone_number)
            calling_search_space = get_calling_search_space(phone_number)

            # Define line appearance object for CUCM
            line = {
                'index': 1,
                'label': full_name,
                'display': full_name,
                'dirn': {
                    'pattern': phone_number,
                    'routePartitionName': {'_value_1': route_partition},
                },
                'ringSetting': 'Use System Default',
                'displayAscii': full_name,
            }

            # Define phone device object for CUCM
            phone = {
                'name': user_id,
                'description': full_name,
                'product': 'Cisco Unified Client Services Framework',
                'class': 'Phone',
                'protocol': 'SIP',
                'devicePoolName': {'_value_1': device_pool},
                'commonPhoneConfigName': {'_value_1': 'Standard Common Phone Profile'},
                'locationName': {'_value_1': 'Hub_None'},
                'phoneTemplateName': {'_value_1': 'Standard Client Services Framework'},
                'callingSearchSpaceName': {'_value_1': calling_search_space},
                'ownerUserName': {'_value_1': user_id},
                'lines': {'line': [line]},
            }

            # Attempt to add phone to CUCM using AXL API
            try:
                response = client.service.addPhone(phone=phone)
                print(f"Device {user_id} added successfully with number {phone_number}.")
            except Exception as e:
                print(f"Failed to add device {user_id}: {e}")
except Exception as e:
    print(f"Failed to read CSV or add devices: {e}")

# --- Call PowerShell Script to Update Active Directory ---

# Define path to the external PowerShell script
ps_script_path = os.path.join(base_path, "adcucm.ps1")

# Call the script with CSV path as parameter
try:
    subprocess.run(
        ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', ps_script_path, '-csvFilePath', user_csv_file],
        check=True
    )
    print("Active Directory update script executed successfully.")
except subprocess.CalledProcessError as e:
    print(f"An error occurred while updating Active Directory: {e}")

# --- Final Message ---

print("Script execution completed.")

