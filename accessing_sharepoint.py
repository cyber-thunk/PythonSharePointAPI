"""Sep 29, 2022 Accessing SharePoint documents via graph.
This is a POC. In prod, will be looking for files in a
directory & filter by file date modified & part of
file name.
"""

import datetime
import json

import msal
import pandas as pd
import requests


date_format = "%Y-%m-%dT%H:%M:%S%z"

def access_sharepoint_data():
    print_header()
    config_data = get_config_data()
    connect_token_dict = get_token_dict(config_data)
    get_sharepoint_data(connect_token_dict, config_data)


def get_sharepoint_data(connect_token_dict:dict, config_data:dict):
    """Access sharepoint data.
        
        Args:
            token_dict (dict): token dict.
            config_data (dict): config dict.
    """
    if not connect_token_dict:
        print(get_sharepoint_data.__name__ + 'Error: No token dict')
    
    auth_string = f"{connect_token_dict.get('token_type')} {connect_token_dict.get('access_token')}"
    
    if 'access_token' in connect_token_dict:
        # Getting group's team site & associated data
        graph_data = requests.get(f"{config_data.get('graph_sites')}",
                                       headers={'Authorization': auth_string})
        graph_data = graph_data.json()

        # Sales dir id
        for drive in graph_data.get('value'):
            if drive['name'] == 'Sales':
                # Get project docs id
                project_docs = requests.get(f"{config_data.get('graph_sites')}/{drive['id']}/root:/US/US02/US0201.3834/Projs",
                                            headers={'Authorization': auth_string})
                project_docs = project_docs.json()
                project_docs_id = project_docs['id']

        # Get project docs items
        project_docs_items = requests.get(f"{config_data.get('graph_sites')}/{drive['id']}/items/{project_docs_id}/children",
                                             headers={'Authorization': auth_string})
        project_docs_items = project_docs_items.json()
        items = list(project_docs_items.get('value'))
        latest_item = generate_sp_doc_data(items)
        sp_file_data = pd.read_excel(latest_item['@microsoft.graph.downloadUrl'], sheet_name='SUMMARY')
        print(sp_file_data)


def generate_sp_doc_data(items:list)->dict:
    """Grab data associated with latest file I'm looking for.
    
    Args:
        items (list): list of sharepoint docs.
        item(dict): sharepoint doc
    """
    for item in items:
        if item['name'][:5] != 'GCTS_':
            items.remove(item)

    for item in items:
        if item['name'][-4:-1] != 'xls':
            items.remove(item)

    latest = None
    last_updated = None
    for item in items:
        item_updated = datetime.datetime.strptime(item['lastModifiedDateTime'], date_format)
        if not last_updated or item_updated > last_updated:
            last_updated = item_updated
            latest = item

    return latest

def get_token_dict(config_data:dict)->dict:
    """Get token from Azure AD.

    Args:
        config_data (dict): config data.

    Returns:
        token_dict (dict): token dict.
    """
    app = msal.ConfidentialClientApplication(
        client_id=config_data.get('sharepoint_client_id'),
        authority=config_data.get('authority'),
        client_credential=config_data.get('sharepoint_client_sct_value')
    )

    result = app.acquire_token_for_client(scopes=config_data.get('sharepoint_scope'))

    if 'access_token' in result:
        token_dict = result
    else:
        token_dict = {}

    return token_dict


def get_config_data()->dict:
    """Get config data from json file.
    
    Returns:
        config_data (dict): config data.
    """
    with open('parameters.json', 'r') as config_file:
        config_data = json.load(config_file)

    return config_data


def print_header():
    print()
    print('---------------------------------')
    print('   ACCESSING SHAREPOINT DATA     ')
    print('  ' + str(datetime.datetime.now()) + ' ')
    print('---------------------------------')
    print()


if __name__ == '__main__':
    access_sharepoint_data()