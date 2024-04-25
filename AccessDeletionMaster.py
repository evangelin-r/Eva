import os
import logging
import configparser
import urllib.parse
import requests
import xlrd

"""
The purpose of the code is to Delete Item Data Security for a given Class.
__author__      = "Mohammed Arif Khan"
__copyright__   = "Copyright 2023"
"""

# Constants for API endpoints
GRP_SEC_BEG = "/fscmRestApi/resources/latest/productManagementDataSecurities?q=ObjectName=Item;ItemClass="
GRP_SEC_END1 = ";OrganizationCode="
GRP_SEC_END2 = ";Principal=Group;Name="
POST_REQ = "/fscmRestApi/resources/latest/productManagementDataSecurities"

def delete_access(sec_responses, class_name, functional_role, org):
    """Delete access for a given class and role."""
    logging.info("Deletion of the access begin")
    for sec_response in sec_responses:
        links = sec_response["links"]
        for link in links:
            if link["rel"] == "self":
                sec_res_link = link["href"]
                try:
                    delete_response = requests.delete(sec_res_link, auth=(user, pwd), headers={"REST-Framework-Version": "1"})
                    delete_response.raise_for_status()
                    logging.info(f"Access deleted for role {functional_role} for Class: {class_name} on Org {org}")
                except requests.exceptions.RequestException as err:
                    logging.error(err)

if __name__ == '__main__':
    print("+++++++++++++++++ Execution of Item Security Class begins +++++++++++++++++")
    # Load configurations from properties file
    prop_file = configparser.RawConfigParser()
    prop_file.read(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Application.properties'))
    environment = prop_file.get("AppDetails", "environment")
    user = prop_file.get("AppDetails", "user")
    pwd = prop_file.get("AppDetails", "password")
    log_file = prop_file.get("AppDetails", "deleteSecLog")
    file_loc = prop_file.get("AppDetails", "delSecTemplate")

    # Configure logging
    logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.INFO, format='%(message)s')

    # Open Excel file
    wb = xlrd.open_workbook(file_loc)
    sheet = wb.sheet_by_index(0)

    # Iterate through Excel rows
    for i in range(1, sheet.nrows):
        item_class = sheet.cell_value(i, 0).strip()
        class_name = urllib.parse.quote(item_class)
        functional_role = sheet.cell_value(i, 1).strip()
        org = sheet.cell_value(i, 2).strip()
        grp_sec_end = GRP_SEC_END1 + org + GRP_SEC_END2

        # Construct API URL
        fun_grp_url = f"{environment}{GRP_SEC_BEG}{class_name};Principal=Group;Name={functional_role};OrganizationCode={org};InheritedFlag=false"

        try:
            logging.info(f"Started executing code for Role: {functional_role} : Class: {item_class}")
            response = requests.get(fun_grp_url, auth=(user, pwd), headers={"REST-Framework-Version": "1"})
            response.raise_for_status()
            jresponse = response.json()
            sec_count = jresponse["count"]
            logging.info(f"Checking Access for Role: {functional_role} | On Class: {item_class}")

            if sec_count == 0:
                logging.info(f"No Action required as Access does not exist for {functional_role} |on Class {item_class}")
 
            if sec_count > 0:
                delete_access(jresponse["items"], class_name, functional_role, org)
                
        except requests.exceptions.RequestException as err:
            logging.error(err)

    logging.info("Execution of Code Completed.")
    print("Execution of Code Completed.")
