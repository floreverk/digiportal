# templatetags/custom_filters.py
from django import template
import base64

register = template.Library()

@register.filter
def base64encode(value):
    """
    Convert the BytesIO object to a base64-encoded string.
    :param value: The BytesIO object containing the image.
    :return: Base64-encoded string representing the image.
    """
    return base64.b64encode(value.getvalue()).decode('utf-8')