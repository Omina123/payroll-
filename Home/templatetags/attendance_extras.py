# Create a file: your_app/templatetags/attendance_extras.py
from django import template
register = template.Library()

@register.filter
def get_item(dictionary, key):
    return dictionary.get(key)