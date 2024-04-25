import os
import logging
import configparser
import urllib.parse
import requests
import xlrd

"""
The purpose of the code is to Create / Update Item Data Security for a given Class.
__author__      = "Mohammed Arif Khan"
__copyright__   = "Copyright 2023"
"""


grpsecbeg = "/fscmRestApi/resources/latest/productManagementDataSecurities?q=ObjectName=Item;ItemClass="
grpsecend1 = ";OrganizationCode="
grpsecend2 = ";Principal=Group;Name="
postreq = "/fscmRestApi/resources/latest/productManagementDataSecurities"

def create_security_access(item_class, functional_role, payload, environment, user, pwd):
    """Create security access for a given item class and functional role.

    Args:
        item_class (str): The item class.
        functional_role (str): The functional role.
        payload (dict): The payload for creating security access.
        environment (str): The environment URL.
        user (str): The user for authentication.
        pwd (str): The password for authentication.
    """
    logging.info("Executing the Create Security Access Block")
    post_url = f"{environment}:443{postreq}"
    try:
        cresponse = requests.post(post_url, json=payload, auth=(user, pwd), headers={"REST-Framework-Version": "1"})
        cresponse.raise_for_status()
        logging.info(f"Access created for role {functional_role} for Class: {item_class}")
    except requests.exceptions.RequestException as err:
        logging.info(err)

def update_security_access(item_class, functional_role, new_payload, links, user, pwd):
    """Update security access for a given item class and functional role.

    Args:
        item_class (str): The item class.
        functional_role (str): The functional role.
        new_payload (dict): The new payload for updating security access.
        links (list): The list of links.
        user (str): The user for authentication.
        pwd (str): The password for authentication.
    """
    logging.info("Executing the Update Security Access block")
    for link in links:
        if link["rel"] == "self":
            sec_res_link = link["href"]
            try:
                logging.info(new_payload)
                logging.info("Update link")
                logging.info(sec_res_link)
                presponse = requests.patch(sec_res_link, json=new_payload, auth=(user, pwd), headers={"REST-Framework-Version": "1"})
                presponse.raise_for_status()
                logging.info(f"Access Updated for Role: {functional_role} : Class: {item_class}")
            except requests.exceptions.RequestException as err:
                logging.info(err)

if __name__ == '__main__':
    print("+++++++++++++++++ Execution of Item Security Class begins +++++++++++++++++")
    prop_file = configparser.RawConfigParser()
    prop_file.read(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Application.properties'))
    environment = prop_file.get("AppDetails", "environment")
    user = prop_file.get("AppDetails", "user")
    pwd = prop_file.get("AppDetails", "password")
    log_file = prop_file.get("AppDetails", "createSecLogFile")
    file_loc = prop_file.get("AppDetails", "createSecTemplate")

    logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.INFO, format='%(message)s')

    wb = xlrd.open_workbook(file_loc)
    sheet = wb.sheet_by_index(0)
    for i in range(1, sheet.nrows):
        item_class = sheet.cell_value(i, 0).strip()
        class_name = urllib.parse.quote(item_class)
        functional_role = sheet.cell_value(i, 1).strip()
        action = sheet.cell_value(i, 2).strip()
        item_eff_actions = sheet.cell_value(i, 3).strip()
        org = sheet.cell_value(i, 4).strip()
        grp_sec_end = grpsecend1 + org + grpsecend2
        # logging.info(str(grp_sec_end))

        fun_grp_url = (f"{environment}{grpsecbeg}{class_name};Principal=Group;Name={functional_role};OrganizationCode={org};"
                       f"InheritedFlag=false"
        )

        logging.info(fun_grp_url)

        try:
            logging.info(f"Started executing code for Role: {functional_role} : Class: {item_class}")
            response = requests.get(fun_grp_url, auth=(user, pwd), headers={"REST-Framework-Version": "1"})
            response.raise_for_status()
            jresponse = response.json()
            sec_count = jresponse["count"]
            logging.info(f"Checking Access for Role: {functional_role} | OnClass: {item_class}")

            if sec_count == 0:
                payload = {
                    "ObjectName": "Item",
                    "InstanceType": "SET",
                    "Principal": "Group",
                    "Name": functional_role,
                    "OrganizationCode": org,
                    "ItemClass": item_class,
                    "Actions": action,
                    "ItemEFFActions": item_eff_actions,
                }
                create_security_access(item_class, functional_role, payload, environment, user, pwd)

            if sec_count >= 1:
                sec_responses = jresponse["items"]
                # logging.info(sec_responses)
                for sec_response in sec_responses:
                    inherited_flag = sec_response["InheritedFlag"]
                    existing_actions = sec_response["Actions"]
                    existing_item_eff_actions = sec_response["ItemEFFActions"]

                    new_payload = {
                        # "ObjectName": "Item",
                        # "InstanceType": "SET",
                        # "Principal": "Group",
                        # "Name": functional_role,
                        # "OrganizationCode": org,
                        # "ItemClass": item_class,
                        "Actions": action,
                        "ItemEFFActions": item_eff_actions
                    }

                    logging.info("Existing Action")
                    links = sec_response["links"]
                    update_security_access(item_class, functional_role, new_payload, links, user, pwd)

        except requests.exceptions.RequestException as err:
            logging.info(err)

    logging.info("Execution of Code Completed.")
    print("Execution of Code Completed.")
