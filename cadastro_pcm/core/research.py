"""No network requests. Research suggestions are entered explicitly by the user."""
from datetime import date
from urllib.parse import urlparse
from .models import Evidence, Item

AVAILABLE = False
STATUS = 'Pesquisa automática indisponível. Consulte fontes autorizadas e registre a URL na edição.'


def apply_suggestion(item: Item, field: str, value, url: str, estimated=False):
    if field not in Item.model_fields or field in ('evidence', 'approved'):
        raise ValueError('Campo não permitido.')
    parsed = urlparse(url)
    if parsed.scheme != 'https' or not parsed.hostname:
        raise ValueError('Informe a URL HTTPS da fonte consultada.')
    data = item.model_dump()
    data[field] = value
    data['approved'] = False
    updated = Item.model_validate(data)
    updated.evidence[field] = Evidence(origin='pesquisa', source=url, consulted=date.today(), estimated=estimated)
    return updated
