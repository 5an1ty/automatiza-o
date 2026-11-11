from datetime import date
from decimal import Decimal
from typing import Literal
from pydantic import BaseModel, Field, field_validator
import re
from .product_text import commercial_line


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
    ncm_not_applicable: bool = False
    warnings: list[str] = Field(default_factory=list)
    # Legacy fields remain readable in existing JSON; new extraction/review does
    # not populate them and document output never adds them automatically.
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
    price_fill_later: bool = False
    authorized_pending_fields: list[Literal['unit_price', 'ncm']] = Field(default_factory=list)
    approved: bool = False
    evidence: dict[str, Evidence] = Field(default_factory=dict)

    @field_validator('ncm')
    @classmethod
    def validate_ncm(cls, value):
        if not value:
            return ''
        if not re.fullmatch(r'(?:\d{8}|\d{4}\.\d{2}\.\d{2})', value):
            raise ValueError('NCM deve conter oito dígitos ou o formato 0000.00.00')
        digits = value.replace('.', '')
        return f'{digits[:4]}.{digits[4:6]}.{digits[6:]}'

    def structural_pending_fields(self):
        missing = []
        if not self.name.strip(): missing.append('Nome do material')
        elif commercial_line(self.name.strip().splitlines()[0]):
            missing.append('Identificação do material (o nome informado é uma condição comercial)')
        if not self.requester.strip(): missing.append('Requisitante')
        return missing

    def deferred_pending_fields(self):
        missing = {}
        if self.unit_price is None and not self.price_fill_later:
            missing['unit_price'] = 'Preço unitário confirmado'
        if not self.ncm and not self.ncm_not_applicable:
            missing['ncm'] = 'NCM fiscal (não informado; preenchimento posterior)'
        return missing

    def pending_fields(self):
        return self.structural_pending_fields() + list(self.deferred_pending_fields().values())

    def blocking_fields(self):
        return self.structural_pending_fields() + [reason for field, reason in self.deferred_pending_fields().items()
                                                   if field not in self.authorized_pending_fields]

    def approve(self, allow_incomplete=False):
        missing = self.structural_pending_fields() if allow_incomplete else self.pending_fields()
        if missing:
            raise ValueError('Campos pendentes: ' + '; '.join(missing) + '.')
        self.authorized_pending_fields = list(self.deferred_pending_fields()) if allow_incomplete else []
        self.approved = True

    def require_ready(self):
        if not self.approved:
            raise ValueError('Revise e aprove o item antes de gerar.')
        if missing := self.blocking_fields():
            raise ValueError('Campos pendentes: ' + '; '.join(missing) + '.')


def decimal_br(text):
    value = re.sub(r'\s+', '', str(text).strip().replace('R$', ''))
    if ',' in value:
        value = value.replace('.', '').replace(',', '.')
    return Decimal(value)


def money_br(value):
    if value is None:
        return 'Pendente'
    return 'R$ ' + f'{value:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
