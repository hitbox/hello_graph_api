import argparse
import configparser
import json
import logging.config
import math
import re

from pprint import pprint

import msal
import requests

try:
    import marshmallow
except ImportError:
    marshmallow = None
else:
    from graphschema import MessageSchema

config_scopes_re = re.compile('^scopes\d+')

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
        password = sec['password'],
        endpoint = sec['endpoint'],
    )
    # assemble numbered scopes names into a list
    config['scopes'] = [sec[key] for key in sec if config_scopes_re.match(key)]
    return config

def graph_get(url, access_token):
    """
    Convenience for request to graph api.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
    }
    logger = logging.getLogger('hello_graph_api')
    logger.debug(url)
    response = requests.get(url, headers=headers)
    return response

def process(config, endpoint, access_token):
    response = graph_get(endpoint, access_token)
    data = response.json()
    return data

def hello_graph_api(config, output=None, limit_next=1):
    logger = logging.getLogger('hello_graph_api')
    if limit_next is None:
        # follow all nextLinks
        limit_next = math.inf

    app = msal.ConfidentialClientApplication(
        client_id = config['client_id'],
        authority = config['authority'],
        client_credential = config['secret'],
    )

    # try to get access token
    accounts = app.get_accounts(username=config['username'])
    for account in accounts:
        logger.debug('acquire_token_silent(%r, account=%r)', scopes, account)
        scopes = config['scopes']
        token_result = app.acquire_token_silent(scopes, account=account)
        if token_result:
            break
    else:
        logger.debug(
            'acquire_token_by_username_password(%r, ***, scopes=%r)',
            config['username'],
            config['scopes']
        )
        token_result = app.acquire_token_by_username_password(
            config['username'],
            config['password'],
            scopes = config['scopes'],
        )
        if not token_result:
            logger.debug('acquire_token_for_client(%r)', config['scopes'])
            token_result = app.acquire_token_for_client(config['scopes'])

    access_token = token_result['access_token']

    # get from endpoint until nextLink is not included
    endpoint = config['endpoint']
    messages = []
    while True:
        try:
            data = process(config, endpoint, access_token)
            messages.extend(data['value'])
            if '@odata.nextLink' not in data:
                break
            limit_next -= 1
            if not limit_next:
                break
            endpoint = data['@odata.nextLink']
        except KeyboardInterrupt:
            # break loop and use what we've got
            break

    if marshmallow:
        message_schema = MessageSchema()
        messages = message_schema.load(messages, many=True, partial=True)

    if not output:
        pprint(messages)
        print(f'{len(messages)=}')
    else:
        with open(output, 'w') as output_f:
            json.dump(messages, output_f, indent=4)

def config_has_logging(cp):
    """
    ConfigParser instance has the sections necessary to config logging.
    """
    return set(['loggers', 'handlers', 'formatters']).issubset(cp)

def main():
    """
    Hello world for Microsoft Graph API.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'config',
        nargs = '+',
        help = 'Path to INI config file.'
    )
    parser.add_argument(
        '--output',
        help = 'Output filename for JSON responses.'
    )
    parser.add_argument(
        '--limit',
        type = int,
        help = 'Number of nextLink links to follow, including the first'
               ' request. Defaults to one request.'
    )
    parser.add_argument(
        '--dump',
        action = 'store_true',
        help = 'Dump processed config.'
    )
    args = parser.parse_args()

    cp = configparser.ConfigParser()
    cp.read(args.config)

    if config_has_logging(cp):
        print('Configuring logging')
        logging.config.fileConfig(cp)

    config = process_config(cp)

    if args.dump:
        from pprint import pprint
        pprint(config)
        parser.exit()

    hello_graph_api(
        config,
        output = args.output,
        limit_next = args.limit,
    )

if __name__ == '__main__':
    main()
