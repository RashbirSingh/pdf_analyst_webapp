import re
from django import template
register = template.Library()

@register.filter
def settingsname(value):
    settingsname = str(value).split('/')[-1]
    return re.sub('^[^_]+_', "", settingsname)