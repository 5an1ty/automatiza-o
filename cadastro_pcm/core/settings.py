import json
import os
from datetime import date
from pathlib import Path
from pydantic import BaseModel, Field
from typing import Literal


class Settings(BaseModel):
    requester: str = ''
    area: str = 'MANUTENÇÃO'
    date_policy: Literal['atual', 'fixa'] = 'atual'
    fixed_date: date = date(2026, 1, 1)
    output: str = str(Path.home() / 'Documents' / 'CadastroPCM')
    template: str = ''
    classification: str = 'Material'
    unit: str = 'UNID'
    group: str = 'Material Manutenção'
    purpose: str = 'Item de compra'
    sap: str = ''
    equipment: str = ''
    imported: Literal['Não', 'Sim'] = 'Não'
    agency: str = 'Item de compra'
    lower_ncm: str = 'SIM'
    sheet: str = 'Criação de Material Nacional'
    # Filled only after inspecting the actual corporate workbook.
    authorized_template_sha256: str = ''
    expected_pages: int = Field(default=1, ge=1)


def settings_path():
    return Path(os.environ.get('LOCALAPPDATA', Path.home() / '.config')) / 'CadastroPCM' / 'settings.json'


def load_settings():
    path = settings_path()
    return Settings.model_validate_json(path.read_text(encoding='utf-8')) if path.exists() else Settings()


def save_settings(settings):
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix('.tmp')
    temp.write_text(settings.model_dump_json(indent=2), encoding='utf-8')
    temp.replace(path)
