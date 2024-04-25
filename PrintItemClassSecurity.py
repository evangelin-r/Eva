import os
import logging
import configparser
import urllib.parse
import requests
import xlrd
import csv
import pandas as pd
from shutil import copyfile

"""
The purpose of the code is to Print Item Data Security for a given Class in the property file.
__author__      = "Mohammed Arif Khan"
__copyright__   = "Copyright 2023"
"""

classoutputfield = "?fields=ItemClass&onlyData=true"  # OutputField of RESTAPI Response
org = "/fscmRestApi/resources/latest/inventoryOrganizationsLOV"  # URI for getting Organizations details
orgoutputfield = "?fields=OrganizationCode&onlyData=true"  # QueryParameter to retrieve response for an organization
# Output field for ItemDataSecurity REST API Response.
secOutputfield = "&fields=ObjectName,Principal,Name,ItemClass,OrganizationCode,Actions,ItemEFFActions,InheritedFlag"
# ItemDataSecurityURL
prodSec = "/fscmRestApi/resources/latest/productManagementDataSecurities?q=ObjectName=Item;ItemClass="
orgfilter = ";OrganizationCode="


def get_security_details(classname, securityurl, role_name, org):
    """For a given pair of Class and Organization, The code extracts all the security entries i.e. Roles,
    Actions, EFF Actions and appends it to a list.

    Args:
        classname (str): Input Class Name
        securityurl (str): URL to get security details
        org (str): Organization code

    Returns:
        None
    """
    logging.debug("Printing security for ClassName: " + classname)
    try:
        sec_res = requests.get(securityurl, auth=(user, pwd), headers={"REST-Framework-Version": "1"})
        sec_res.raise_for_status()
        jsec_res = sec_res.json()
        if jsec_res["count"] > 0:
            sec_datas = jsec_res["items"]
            for sec_data in sec_datas:
                logging.info(
                    f"{sec_data['ObjectName']}~{sec_data['Principal']}~{sec_data['ItemClass']}~{sec_data['Name']}~"
                    f"{sec_data['Actions']}~{sec_data['ItemEFFActions']}~{sec_data['OrganizationCode']}~"
                    f"{sec_data['InheritedFlag']}")
        elif jsec_res["count"] == 0:
            logging.info(
                f"No existing Security defined for Role {role_name} on Class {classname} on Organization {org}")
    except requests.exceptions.RequestException as err:
        logging.debug(err)


if __name__ == '__main__':
    print("+++++++++++++++++ Execution of Item Security Class begins +++++++++++++++++")
    prop_file = configparser.RawConfigParser()
    prop_file.read(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Application.properties'))

    environment = prop_file.get("AppDetails", "environment")
    user = prop_file.get("AppDetails", "user")
    pwd = prop_file.get("AppDetails", "password")
    log_file = prop_file.get("AppDetails", "printlogFile")
    current_directory = os.getcwd()
    file_loc = prop_file.get("AppDetails", "printSecTemplate")
    outputfile = prop_file.get("AppDetails", "outputSecFile")

    # Backup the existing log file
    backup_file = log_file + ".backup"
    copyfile(log_file, backup_file)

    # Clear the log file
    open(log_file, 'w').close()

    logging.basicConfig(filename=log_file, encoding='utf-8', level=logging.INFO, format='%(message)s')
    logging.debug(f"{environment}\n{file_loc}\n{user}\n{pwd}\n")
    wb = xlrd.open_workbook(file_loc)
    sheet = wb.sheet_by_index(0)
    # logging.info(
    #     "ObjectName~Principal~ItemClass~Role~Actions~ItemEFFActions~OrganizationCode~InheritedFlag")
    for i in range(1, sheet.nrows):
        item_class = sheet.cell_value(i, 0).strip()
        class_name = urllib.parse.quote(item_class)
        role_name = sheet.cell_value(i, 1).strip()
        org_code = sheet.cell_value(i, 2).strip()
        security_url = (
                f"{environment}{prodSec}{class_name};Principal=Group;Name={role_name};OrganizationCode={org_code};"
                f"&onlyData=true{secOutputfield}"
        )
        get_security_details(class_name, security_url, role_name, org_code)

    # Read log file line by line
    with open(log_file, 'r') as log_file:
        log_lines = log_file.readlines()
    # Split each line based on "~" delimiter and store in a list
    log_data = [line.strip().split('~') for line in log_lines]
    # Convert list of lists to DataFrame
    df = pd.DataFrame(log_data, columns=["ObjectName", "Principal", "ItemClass", "Role", "Actions", "ItemEFFActions",
                                         "OrganizationCode", "InheritedFlag"])
    # Write DataFrame to Excel file
    df.to_excel(outputfile, index=False)
