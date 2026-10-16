from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    if isinstance(dictionary, dict):
        return dictionary.get(key, "")
    return ""

@register.filter
def format_phone(val):
    if not val:
        return "-"
    s = str(val).strip()
    if len(s) == 10:
        return f"+91 {s[:5]} {s[5:]}"
    return s

@register.filter
def phone_status_badge(status):
    """Returns HTML badge class for phone status."""
    st = str(status).lower() if status else 'missing'
    if st in ('valid',):
        return 'badge-valid'
    elif st in ('cleaned',):
        return 'badge bg-info-subtle text-info border border-info-subtle'
    elif st in ('multiple_numbers_review',):
        return 'badge bg-warning-subtle text-dark border border-warning'
    elif st in ('invalid', 'flagged'):
        return 'badge badge-flagged'
    else:  # missing, blank
        return 'badge badge-blank'


@register.simple_tag(takes_context=True)
def url_replace(context, **kwargs):
    query = context['request'].GET.copy()
    for key, value in kwargs.items():
        if value is None or value == '':
            query.pop(key, None)
        else:
            query[key] = value
    return query.urlencode()

