import argparse
import base64
import codecs
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
    from marshmallow import fields

    class Base64ContentField(fields.Field):

        def _deserialize(self, string_of_bytes, attr, data, **kwargs):
            try:
                return codecs.decode(base64.b64decode(string_of_bytes))
            except TypeError:
                raise marshmallow.ValidationError(
                    'Base 64 content field must be string or bytes.')

    class BodySchema(marshmallow.Schema):

        content = fields.String()
        content_type = fields.String(data_key='contentType')


    class EmailAddressSchema(marshmallow.Schema):

        address = fields.String()
        name = fields.String()


    class SenderSchema(marshmallow.Schema):

        email_address = fields.Nested(EmailAddressSchema, data_key='emailAddress')


    class AttachmentSchema(marshmallow.Schema):

        class Meta:
            unknown = marshmallow.EXCLUDE


        content_type = fields.String(data_key='contentType')
        content = Base64ContentField(data_key='contentBytes')
        id = fields.String()
        is_inline = fields.Boolean(data_key='isInline')
        last_modified_datetime = fields.DateTime(data_key='lastModifiedDateTime')
        name = fields.String()
        size = fields.Integer()


    class MessageSchema(marshmallow.Schema):

        class Meta:
            unknown = marshmallow.EXCLUDE


        sender = fields.Nested(SenderSchema)
        subject = fields.String()
        received_datetime = fields.DateTime(data_key='receivedDateTime')
        body = fields.Nested(BodySchema)

        attachments = fields.List(
            fields.Nested(AttachmentSchema)
        )


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

    pprint(messages)
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
