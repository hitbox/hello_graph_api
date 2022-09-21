import base64
import codecs

import marshmallow

from marshmallow import fields

class Base64ContentField(fields.Field):
    """
    A string (not bytes) of base64 encoded "bytes".
    https://learn.microsoft.com/en-us/graph/api/resources/fileattachment?view=graph-rest-1.0#properties
    See contentBytes
    """

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
        # ignore metadata like links to the next page and other properties
        # not needed.
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
        # ignore metadata like links to next page and other uneeded properties.
        unknown = marshmallow.EXCLUDE


    sender = fields.Nested(SenderSchema)
    subject = fields.String()
    received_datetime = fields.DateTime(data_key='receivedDateTime')
    body = fields.Nested(BodySchema)

    attachments = fields.List(
        fields.Nested(AttachmentSchema)
    )
