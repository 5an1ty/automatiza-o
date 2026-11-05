"""Portable unit tests never launch the user's Excel or use corporate files.
Native Excel is exercised through explicit fakes; real COM is a local acceptance check.
"""
import pytest


@pytest.fixture(autouse=True)
def portable_excel_backend(monkeypatch):
    from cadastro_pcm.core import excel_writer
    monkeypatch.setattr(excel_writer, 'use_native_excel', lambda: False)
