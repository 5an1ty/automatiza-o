from datetime import date
from decimal import Decimal
from typing import Literal
from pydantic import BaseModel, Field, field_validator
import re


class Evidence(BaseModel):
    origin: Literal['pdf', 'regra', 'manual', 'pesquisa', 'pendente']
    source: str = ''
    excerpt: str = ''
    consulted: date | None = None
    estimated: bool = False


class Item(BaseModel):
    name: str = ''
    description: str = ''
    quantity: Decimal | None = Field(default=None, gt=0)
    unit_price: Decimal | None = Field(default=None, ge=0)
    total_price: Decimal | None = Field(default=None, ge=0)
    ncm: str = ''
    manufacturer: str = ''
    manufacturer_code: str = ''
    classification: str = 'Material'
    unit: str = 'UNID'
    group: str = 'Material Manutenção'
    purpose: str = 'Item de compra'
    sap: str = ''
    equipment: str = ''
    imported: str = 'Não'
    agency: str = 'Item de compra'
    lower_ncm: str = 'SIM'
    requester: str = ''
    area: str = 'MANUTENÇÃO'
    request_date: date = Field(default_factory=date.today)
    approved: bool = False
    evidence: dict[str, Evidence] = Field(default_factory=dict)

    @field_validator('ncm')
    @classmethod
    def validate_ncm(cls, value):
        if not value:
            return ''
        if not re.fullmatch(r'(?:\d{8}|\d{4}\.\d{2}\.\d{2})', value):
            raise ValueError('NCM deve conter oito dígitos ou o formato 0000.00.00')
        return value.replace('.', '')

    def require_ready(self):
        if not self.approved:
            raise ValueError('Revise e aprove o item antes de gerar.')
        if not self.name.strip() or not self.requester.strip():
            raise ValueError('Nome e requisitante são obrigatórios.')
        if self.unit_price is None:
            raise ValueError('Preço unitário pendente. Informe um valor confirmado.')


def decimal_br(text):
    value = str(text).strip().replace('R$', '').replace(' ', '')
    if ',' in value:
        value = value.replace('.', '').replace(',', '.')
    return Decimal(value)


def money_br(value):
    if value is None:
        return 'Pendente'
    return 'R$ ' + f'{value:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
