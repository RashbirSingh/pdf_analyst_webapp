import os, time
from django import template
register = template.Library()

@register.filter
def filetime(value):
    return time.ctime(os.path.getctime(value.file.name))