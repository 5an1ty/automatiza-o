"""Correct stored serial plus a localized Excel renderer, never a real COM claim."""
from contextlib import contextmanager
from datetime import date, datetime, timezone
import json
from types import SimpleNamespace

import pytest
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, from_excel, to_excel

from cadastro_pcm.core import excel_writer
from cadastro_pcm.core.excel_date import (DATE_FORMAT, PDF_DATE_FORMAT, LOCAL_DATE_FORMAT,
    DateValidationError, check_com_date, inspect_com_date)
from cadastro_pcm.core.template_reader import read_template
from test_prompt04 import complete_item, extended_template

EXPECTED = date(2026, 11, 6)
EPOCH = CALENDAR_WINDOWS_1900


class LocalizedCell:
    Width = 134.25; ColumnWidth = 23.14; MergeCells = False

    def __init__(self, local_works=True, invariant_works=True):
        self.Value2 = 46332.0
        self._format = DATE_FORMAT; self._local = DATE_FORMAT; self._text = '06/11/yyyy'
        self.local_works = local_works; self.invariant_works = invariant_works
        self.changes = []
    @property
    def Value(self): return from_excel(self.Value2).replace(tzinfo=timezone.utc)
    @property
    def NumberFormat(self): return self._format
    @NumberFormat.setter
    def NumberFormat(self, value):
        self.changes.append(('NumberFormat', value))
        if not self.invariant_works: raise RuntimeError('Synthetic refused invariant format')
        self._format = value
        self._text = from_excel(self.Value2).strftime('%d/%m/%Y')
    @property
    def NumberFormatLocal(self): return self._local
    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.changes.append(('NumberFormatLocal', value))
        if not self.local_works: raise RuntimeError('Synthetic refused local format')
        self._local = value; self._format = 'dd"/"mm"/"yyyy'
        self._text = from_excel(self.Value2).strftime('%d/%m/%Y')
    @property
    def Text(self): return self._text


@pytest.mark.parametrize('text', ['06/11/yyyy', 'dd/11/yyyy', '06/mm/2026', '06/11/aaaa'])
def test_literal_mask_is_diagnosable_without_exposing_arbitrary_text(text):
    cell = LocalizedCell(); cell._text = text
    report = inspect_com_date(cell, EXPECTED, EPOCH)
    assert report['date_ok'] and report['format_ok'] and not report['display_ok']
    assert report['text'] == text and report['text_kind'] == 'literal_date_mask'
    cell._text = 'PESSOA CONFIDENCIAL dd/11/yyyy'
    assert 'PESSOA' not in json.dumps(inspect_com_date(cell, EXPECTED, EPOCH))


@pytest.mark.parametrize('local_works', [True, False])
def test_verified_local_format_then_invariant_fallback(tmp_path, local_works):
    cell = LocalizedCell(local_works=local_works)
    report = check_com_date(cell, EXPECTED, EPOCH, phase='antes de salvar', report_dir=tmp_path, repair=True, exact=True)
    assert cell.changes[0] == ('NumberFormatLocal', LOCAL_DATE_FORMAT)
    assert report['display'] == 'ok' and report['date_ok'] and report['format_ok']
    assert cell.Value2 == 46332.0 and cell.Width == 134.25 and cell.ColumnWidth == 23.14
    if local_works: assert len(cell.changes) == 1
    else: assert cell.changes[-1] == ('NumberFormat', PDF_DATE_FORMAT)
    saved = json.loads(next(tmp_path.glob('diagnostico_data_*.json')).read_text())
    assert saved['attempts'][0]['text'] == '06/11/yyyy'
    assert saved['attempts'][-1]['text'] == '06/11/2026'
    assert saved['date_ok'] and saved['format_ok'] and saved['display_ok']
    # Reopening/reading does not require a new repair once rendering is correct.
    count = len(cell.changes)
    check_com_date(cell, EXPECTED, EPOCH, phase='após reabrir', report_dir=tmp_path)
    assert len(cell.changes) == count


def test_unverified_local_repair_still_blocks_generation(tmp_path):
    cell = LocalizedCell(local_works=False, invariant_works=False)
    with pytest.raises(DateValidationError) as error:
        check_com_date(cell, EXPECTED, EPOCH, phase='antes de salvar', report_dir=tmp_path, repair=True)
    assert error.value.diagnostic['date_ok'] and not error.value.diagnostic['display_ok']
    assert error.value.diagnostic['attempts'][-1]['text'] == '06/11/yyyy'
    assert len(cell.changes) == 2 and cell.Value2 == 46332.0


def test_wrong_serial_is_never_repaired_as_presentation(tmp_path):
    cell = LocalizedCell(); cell.Value2 += 1
    with pytest.raises(DateValidationError, match='data armazenada'):
        check_com_date(cell, EXPECTED, EPOCH, phase='antes de salvar', report_dir=tmp_path, repair=True)
    assert not cell.changes


def test_native_copy_reopens_after_verified_local_format(extended_template, tmp_path, monkeypatch):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path / 'local'))
    calls, local_changes = [], []
    original = extended_template.read_bytes()
    @contextmanager
    def session(path, read_only):
        calls.append(read_only)
        wb = read_template(path); sheet = wb.worksheets[0]
        class Cell:
            HasFormula = False; MergeCells = False; Locked = False
            Width = 134.25; ColumnWidth = 23.14
            def __init__(self, address): self.cell = sheet[address]
            @property
            def Value(self): return self.cell.value
            @Value.setter
            def Value(self, value): self.cell.value = value
            @property
            def Value2(self):
                value = self.cell.value
                return to_excel(value, wb.epoch) if isinstance(value, datetime) else value
            @Value2.setter
            def Value2(self, value): self.cell.value = value
            @property
            def NumberFormat(self): return self.cell.number_format
            @NumberFormat.setter
            def NumberFormat(self, value): self.cell.number_format = value
            @property
            def NumberFormatLocal(self): return self.cell.number_format
            @NumberFormatLocal.setter
            def NumberFormatLocal(self, value):
                local_changes.append(value); self.cell.number_format = 'dd"/"mm"/"yyyy'
            @property
            def Text(self):
                return '06/11/yyyy' if self.cell.number_format == DATE_FORMAT else '06/11/2026'
        ws = SimpleNamespace(ProtectContents=False, Range=Cell,
                             Cells=SimpleNamespace(SpecialCells=lambda _: object()),
                             Application=SimpleNamespace(Intersect=lambda a, b: None))
        yield None, SimpleNamespace(ReadOnly=read_only, Worksheets=lambda _: ws, Save=lambda: wb.save(path))
        wb.close()
    monkeypatch.setattr(excel_writer, 'excel_session', session)
    monkeypatch.setattr(excel_writer, 'use_native_excel', lambda: True)
    item = complete_item(); item.request_date = EXPECTED
    output = tmp_path / 'copy.xlsx'
    excel_writer.write_excel(extended_template, output, item, 'Criação de Material Nacional', excel_writer.fingerprint(extended_template))
    assert calls == [False, True] and local_changes == [LOCAL_DATE_FORMAT]
    assert extended_template.read_bytes() == original
    wb = read_template(output)
    assert wb.worksheets[0]['E8'].value == datetime(2026, 11, 6)
    assert wb.worksheets[0]['C16'].value is None
    assert wb.worksheets[0].protection.sheet
    wb.close()
