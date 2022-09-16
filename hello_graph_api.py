import argparse
import configparser
import re

from pprint import pprint

import msal
import requests

scopes_re = re.compile('^scopes\d+')

def process_config(cp):
    """
    Raise for missing required names and process for data types.
    """
    sec = cp['hello_graph_api']
    config = dict(
        client_id = sec['client_id'],
        tenant_id = sec['tenant_id'],
        authority = sec['authority'],
        secret = sec['secret'],
        username = sec['username'],
    )
    # assemble numbered scopes names into a list
    config['scopes'] = [sec[key] for key in sec if scopes_re.match(key)]
    return config

def graph_get(url, access_token):
    """
    Convenience for request to graph api.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
    }
    response = requests.get(url, headers=headers)
    return response

def hello_graph_api(config):
    # working out how to use ms graph api
    app = msal.ConfidentialClientApplication(
        client_id = config['client_id'],
        authority = config['authority'],
        client_credential = config['secret'],
    )

    result = app.acquire_token_for_client(config['scopes'])
    access_token = result['access_token']

    username = config['username']
    url = f'https://graph.microsoft.com/v1.0/users/{username}/messages'
    response = graph_get(url, access_token)
    pprint(response.json())

def main():
    """
    Hello world for Microsoft Graph API.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('config', nargs='+')
    parser.add_argument('--dump', action='store_true', help='Dump processed config.')
    args = parser.parse_args()

    cp = configparser.ConfigParser()
    cp.read(args.config)
    config = process_config(cp)

    if args.dump:
        from pprint import pprint
        pprint(config)
        parser.exit()

    hello_graph_api(config)

if __name__ == '__main__':
    main()
