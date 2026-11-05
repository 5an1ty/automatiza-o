from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
import json
import sys
import pytest
import pymupdf
from openpyxl import Workbook, load_workbook
from cadastro_pcm.core.models import Item
from cadastro_pcm.core.pdf_reader import Page, read_pdf
from cadastro_pcm.core.item_extractor import extract_items
from cadastro_pcm.core.print_layout import print_area, configure_page
from cadastro_pcm.core import pdf_exporter, excel_writer, settings, diagnostics


def pdf_at(path, texts):
    with pymupdf.open() as doc:
        for text in texts:
            page = doc.new_page()
            if text: page.insert_text((40,40), text)
        doc.save(path)


def test_multiline_and_missing_fields():
    items = extract_items([Page('documento', 1, 'Produto: Gaxeta 16\nDescrição técnica: Vedação\npara eixo de bomba\nQuantidade: 9\nValor total: 2.151,00\nProduto: Gaxeta 18\nPreço unitário: 250,00', [])])
    assert len(items) == 2
    assert items[0].description == 'Vedação\npara eixo de bomba'
    assert items[0].unit_price is None and items[0].ncm == ''
    assert items[1].name == 'Gaxeta 18' and items[1].quantity is None
    assert items[1].total_price is None


def test_duplicate_detection_never_merges():
    pages = [Page('a', 1, 'Produto: Anel\nPreço unitário: 10,00', []), Page('b', 2, 'Produto: Anel\nPreço unitário: 11,00', []), Page('b', 3, 'Produto: Anel maior\nPreço unitário: 12,00', [])]
    items = extract_items(pages)
    assert len(items) == 3
    assert items[0].warnings and items[1].warnings
    assert not items[2].warnings
    assert items[0].unit_price == 10 and items[1].unit_price == 11


def test_ncm_readiness():
    item = Item(name='Serviço', requester='Responsável', approved=True, unit_price=1)
    with pytest.raises(ValueError, match='NCM'): item.require_ready()
    item.ncm_not_applicable = True; item.require_ready()
    assert Item(ncm='4016.93.00').ncm == Item(ncm='40169300').ncm


def test_print_bounds():
    wb = Workbook(); ws = wb.active; ws.title = 'Criação de Material Nacional'
    ws['A33'] = 'Controle de alterações'
    assert print_area(ws) == 'A1:E33'
    ws['A34'] = 'Assinatura fora do intervalo'
    with pytest.raises(ValueError, match='fora'): print_area(ws)
    del ws['A34']; ws.merge_cells('A14:F14')
    with pytest.raises(ValueError, match='mesclada'): print_area(ws)


def test_page_configuration():
    ws = SimpleNamespace(PageSetup=SimpleNamespace())
    configure_page(ws, 'A1:E34')
    assert ws.PageSetup.PaperSize == 9 and ws.PageSetup.Orientation == 1
    assert ws.PageSetup.FitToPagesWide == 1
    assert ws.PageSetup.FitToPagesTall == 1


def test_pdf_additional_blank_and_missing_name(tmp_path):
    path = tmp_path/'saida.pdf'
    pdf_at(path, ['Gaxeta 16', 'Extra'])
    with pytest.raises(ValueError, match='2 páginas'): pdf_exporter.validate_pdf(path, 'Gaxeta 16')
    pdf_exporter.validate_pdf(path, 'Gaxeta 16', 2)
    path.unlink(); pdf_at(path, ['Gaxeta 16', ''])
    with pytest.raises(ValueError, match='sem texto'): pdf_exporter.validate_pdf(path, 'Gaxeta 16', 2)
    path.unlink(); pdf_at(path, ['Outro produto'])
    with pytest.raises(ValueError, match='Nome'): pdf_exporter.validate_pdf(path, 'Gaxeta 16')


def test_pdf_export_failure_cleans_partial(monkeypatch, tmp_path):
    template = tmp_path/'modelo.xlsx'; wb = Workbook(); ws = wb.active; ws.title='Criação de Material Nacional'; ws['A33']='Controle'; wb.save(template)
    output = tmp_path/'saida.pdf'
    class FakeSheet:
        PageSetup = SimpleNamespace(LeftMargin=36, RightMargin=36, TopMargin=36, BottomMargin=36)
        def Range(self, address): return SimpleNamespace(Value=None, Width=500, Height=700, Text='Controle')
        def ResetAllPageBreaks(self): pass
        def ExportAsFixedFormat(self, *args):
            output.write_bytes(b'incompleto')
            raise RuntimeError('Falha simulada de exportação')
    @contextmanager
    def fake_session(*args, **kwargs):
        yield None, SimpleNamespace(Worksheets=lambda name: FakeSheet())
    monkeypatch.setattr(pdf_exporter, 'excel_session', fake_session)
    monkeypatch.setattr(pdf_exporter.sys, 'platform', 'win32')
    before = template.read_bytes()
    with pytest.raises(RuntimeError, match='simulada'):
        pdf_exporter.export_pdf(template, output, ws.title, 'Gaxeta')
    assert not output.exists() and template.read_bytes() == before


def test_settings_persistence(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path))
    data = settings.Settings(requester='Nome local', date_policy='fixa', fixed_date='2026-11-05', output=str(tmp_path/'saida'))
    settings.save_settings(data)
    assert settings.load_settings() == data
    assert not settings.settings_path().with_suffix('.tmp').exists()


def test_diagnostics_report_privacy(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, 'platform', 'linux')
    data = settings.Settings(requester='NOME CONFIDENCIAL', template=str(tmp_path/'SEGREDO.xlsx'), output=str(tmp_path/'saida'))
    report = diagnostics.diagnose(data)
    text = json.dumps(report, ensure_ascii=False)
    assert 'NOME CONFIDENCIAL' not in text and 'SEGREDO' not in text and str(tmp_path) not in text
    assert any(check['status']=='não executado' for check in report['verificacoes'])
    assert not list((tmp_path/'saida').iterdir())
    diagnostics.save_report(report, tmp_path/'diagnostico.json')
    assert json.loads((tmp_path/'diagnostico.json').read_text()) == report


def test_com_session_releases_after_open_failure(monkeypatch):
    from cadastro_pcm.core import excel_runtime
    calls=[]
    pythoncom=SimpleNamespace(CoInitialize=lambda:calls.append('init'), CoUninitialize=lambda:calls.append('uninit'))
    class Books:
        def Open(self,*args,**kwargs): raise RuntimeError('open failed')
    excel=SimpleNamespace(Workbooks=Books(), Quit=lambda:calls.append('quit'))
    client=SimpleNamespace(DispatchEx=lambda name: excel)
    monkeypatch.setitem(sys.modules,'pythoncom',pythoncom)
    monkeypatch.setitem(sys.modules,'win32com',SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules,'win32com.client',client)
    monkeypatch.setattr(sys,'platform','win32')
    with pytest.raises(RuntimeError,match='open failed'):
        with excel_runtime.excel_session('modelo.xlsx'): pass
    assert calls==['init','quit','uninit']


def test_com_writer_failure_removes_copy(monkeypatch, tmp_path):
    source = tmp_path/'modelo.xlsx'; source.write_bytes(b'original')
    dest = tmp_path/'copy.xlsx'
    @contextmanager
    def failed_session(*args,**kwargs):
        raise RuntimeError('Excel indisponível')
        yield
    monkeypatch.setattr(excel_writer,'excel_session',failed_session)
    with pytest.raises(RuntimeError):
        excel_writer.write_excel_com(source,dest,{},'Cadastro',None)
    assert not dest.exists() and source.read_bytes()==b'original'


def test_multiple_pdf_pages_no_shared_values(tmp_path):
    path=tmp_path/'varias.pdf'
    pdf_at(path,['Produto: Gaxeta 16\nPreço unitário: 239,00\nNCM: 4016.93.00','Produto: Gaxeta 18\nQuantidade: 2'])
    items=extract_items(read_pdf(path))
    assert len(items)==2 and items[1].unit_price is None and items[1].ncm==''
    assert 'página 2' in items[1].evidence['name'].source


def test_native_com_validation_rejects_value(monkeypatch, tmp_path):
    source=tmp_path/'original.xlsx'; source.write_bytes(b'nao alterar')
    target=tmp_path/'copia.xlsx'
    cell=SimpleNamespace(Value=None,HasFormula=False,MergeCells=False,Validation=SimpleNamespace(Value=False))
    ws=SimpleNamespace(ProtectContents=False)
    ws.Range=lambda address: SimpleNamespace(Value=None) if address=='C16' else cell
    ws.Cells=SimpleNamespace(SpecialCells=lambda kind: object())
    ws.Application=SimpleNamespace(Intersect=lambda a,b: object())
    calls=[]
    book=SimpleNamespace(ReadOnly=False,Worksheets=lambda name:ws,Save=lambda:calls.append('save'))
    @contextmanager
    def fake_session(*args,**kwargs):
        yield None,book
    monkeypatch.setattr(excel_writer,'excel_session',fake_session)
    with pytest.raises(ValueError,match='rejeitado'):
        excel_writer.write_excel_com(source,target,{'A16':'Inválido'},'Cadastro',None)
    assert not target.exists() and not calls and source.read_bytes()==b'nao alterar'


def test_type_not_empty_blocks_write(tmp_path):
    wb=Workbook(); ws=wb.active; ws.title='Criação de Material Nacional'; ws['C16']='Não apagar'
    path=tmp_path/'modelo.xlsx'; wb.save(path)
    item=Item(name='Material',requester='Revisor',unit_price=1,ncm='40169300',approved=True)
    with pytest.raises(ValueError,match='Tipo'):
        excel_writer.write_excel(path,tmp_path/'saida.xlsx',item,ws.title,excel_writer.fingerprint(path))
    assert not (tmp_path/'saida.xlsx').exists()


def test_word_aligned_columns_pdf(tmp_path):
    path=tmp_path/'colunas.pdf'
    with pymupdf.open() as doc:
        page=doc.new_page()
        rows=[['Descrição','Quantidade','Valor unitário','NCM'],['Gaxeta 16','9','239,00','4016.93.00'],['Gaxeta 18','2','250,00','4016.93.00']]
        for row,values in enumerate(rows):
            for x,value in zip([30,200,310,430],values):
                page.insert_text((x,50+row*25),value,fontsize=10)
        doc.save(path)
    items=extract_items(read_pdf(path))
    assert len(items)==2
    assert items[0].name=='Gaxeta 16' and items[0].unit_price==239 and items[0].quantity==9
    assert items[1].name=='Gaxeta 18' and items[1].unit_price==250
