from cadastro_pcm.core.records import material_record
from decimal import Decimal
from pathlib import Path
import pytest
from cadastro_pcm.core.material_list import parse_material_list, SOURCE
from cadastro_pcm.core.models import Item
from cadastro_pcm.core import generation
from cadastro_pcm.core.settings import Settings
from cadastro_pcm.core.excel_writer import fingerprint
from test_workflow import template


@pytest.mark.parametrize('code', ['8481.80.95', '84818095', 'NCM: 8481.80.95', 'NCM 84818095'])
def test_one_material_ncm_formats(code):
    items, notices = parse_material_list(f'Válvula esfera\n{code}\nR$ 2.500,00')
    assert not notices and len(items) == 1
    assert items[0].ncm == '8481.80.95' and items[0].unit_price == Decimal('2500.00')
    assert not items[0].approved and items[0].evidence['name'].source == SOURCE


@pytest.mark.parametrize('text', [
    'Produto A\nNCM: 84818095\nR$ 120,00\n\nProduto B\nNCM 73072900\n45,00',
    'Produto A | 84818095 | R$ 120,00\nProduto B | 73072900 | 45,00',
    'Cadastro\tNCM\tPreço unitário\nProduto A\t84818095\t120,00\nProduto B\t73072900\t45,00',
    '| Cadastro | NCM | Preço unitário |\n|---|:---:|---:|\n| Produto A | 84818095 | 120,00 |\n| Produto B | 73072900 | 45,00 |',
    '# Lista de materiais\n\n**Produto A**\n\n**NCM: 84818095**\nR$ 120,00\n---\nProduto B\nNCM 73072900\n45,00',
])
def test_formats_multiple(text):
    items, notices = parse_material_list(text)
    assert not notices and [i.name for i in items] == ['Produto A', 'Produto B']
    assert [i.unit_price for i in items] == [Decimal('120'), Decimal('45')]
    assert [i.ncm for i in items] == ['8481.80.95', '7307.29.00']


def test_nineteen_synthetic_records():
    # Synthetic reference only: the user's corporate list was not supplied.
    text = '\n\n'.join(f'Produto {n} 90° AISI 304\nNCM: 84818095\nR$ {n},00' for n in range(1, 20))
    items, notices = parse_material_list(text)
    assert not notices and len(items) == 19
    for n, item in enumerate(items, 1):
        assert item.name == f'Produto {n} 90° AISI 304'
        assert item.unit_price == Decimal(n) and item.ncm == '8481.80.95'


def test_technical_description_preserved():
    name = 'Válvula, rosca 1¼" e 1½", 90°; 600 × 310 × 100 mm AISI 304 HC-STA-B06-CHWS-1TTM20-ELC-AL ' + 'descrição longa ' * 35
    items, _ = parse_material_list(name + '\nNCM 84818095\nR$ 250,00')
    assert len(items) == 1 and items[0].name == name.strip() and items[0].description == name.strip()


@pytest.mark.parametrize('text,ncm,amount', [
    ('Produto A\nNCM 84818095', '8481.80.95', None),
    ('Produto A\nR$ 250,00', '', Decimal('250')),
    ('Produto A', '', None),
    ('Produto A | | 250.00', '', Decimal('250')),
    ('Produto A\nPreço: 250', '', Decimal('250')),
    ('Produto A\nPreço unitário: 250.00', '', Decimal('250')),
])
def test_missing_and_explicit_prices(text, ncm, amount):
    items, _ = parse_material_list(text)
    assert len(items) == 1 and items[0].ncm == ncm and items[0].unit_price == amount


def test_no_inference_from_product_numbers():
    items, _ = parse_material_list('Peça 84818095 modelo 250.00\nConexão 600 × 310 × 100 mm')
    assert len(items) == 2
    assert all(not i.ncm and i.unit_price is None for i in items)


def test_duplicates_ambiguity_and_no_cross_assignment():
    items, notices = parse_material_list('R$ 999,00\nProduto A\nNCM 84818095\nProduto B\nR$ 45,00\nProduto A\nR$ 10,00')
    assert notices and len(items) == 3
    assert items[0].unit_price is None and items[1].ncm == ''
    assert items[1].unit_price == Decimal('45')
    assert 'duplicata' in ' '.join(items[0].warnings)
    items, _ = parse_material_list('Produto\nNCM 84818095\nNCM 73072900\nR$ 45,00\nR$ 55,00')
    assert not items[0].ncm and items[0].unit_price is None and items[0].warnings


def test_repeated_headers_and_reordered_columns():
    items, _ = parse_material_list('Preço unitário\tCadastro\tNCM\n120,00\tProduto A\t84818095\nPreço unitário\tCadastro\tNCM\n45,00\tProduto B\t73072900')
    assert [i.unit_price for i in items] == [Decimal('120'), Decimal('45')]


@pytest.mark.parametrize('count', [2, 19])
def test_no_pdf_batch_approval_generation_preserves_template(template, tmp_path, monkeypatch, count):
    items, _ = parse_material_list('Produto A | 84818095 | 120,00\nProduto B | |')
    if count == 19:
        items, _ = parse_material_list('\n'.join(f'Produto {n} | 84818095 | {n},00' for n in range(1, 20)))
    before = template.read_bytes()
    settings = Settings(requester='Revisor', template=str(template), output=str(tmp_path / 'out'), authorized_template_sha256=fingerprint(template))
    for item in items:
        item.requester = settings.requester
        with pytest.raises(ValueError): item.require_ready()
        item.approve(allow_incomplete=True)
    def export(source, target, sheet, name, pages):
        import pymupdf
        with pymupdf.open() as doc:
            page = doc.new_page(); page.insert_text((40, 40), name); doc.save(target)
    monkeypatch.setattr(generation, 'export_pdf', export)
    _, results = generation.generate_batch(list(enumerate(items)), settings)
    assert len(results) == count and all(r.status == 'sucesso' for r in results)
    for result in results:
        assert generation.material_pdf(result.folder).is_file()
        assert generation.material_pdf(result.folder, shared=True).parent == Path(settings.output) / 'lote pdf'
        assert (result.folder / 'cadastro.xlsx').is_file()
        saved = Item.model_validate_json(material_record(result.folder).read_text())
        assert saved.evidence['name'].origin == 'manual'
    assert template.read_bytes() == before


def test_dialog_edit_confirm_invalidate_cancel_and_window_import(monkeypatch):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    from PySide6.QtWidgets import QApplication, QDialog
    from cadastro_pcm.interface.material_list_dialog import MaterialListDialog
    from cadastro_pcm.interface import main_window
    app = QApplication.instance() or QApplication([])
    dialog = MaterialListDialog()
    dialog.text.setPlainText('Produto A | 84818095 | 120,00')
    dialog.analyze(); assert dialog.table.rowCount() == 1 and dialog.add_button.isEnabled()
    dialog.table.item(0, 0).setText('Produto revisado 1½"')
    dialog.table.item(0, 2).setText('250,00')
    dialog.confirm()
    assert dialog.result_items[0].name == 'Produto revisado 1½"'
    assert dialog.result_items[0].unit_price == Decimal('250')
    dialog.text.setPlainText('Outro'); assert not dialog.add_button.isEnabled()
    monkeypatch.setattr(main_window, 'load_settings', lambda: Settings(requester='Revisor'))
    window = main_window.Window(); window.items = [Item(name='Existente')]
    def accept(self):
        self.text.setPlainText('Produto A | 84818095 | 120,00'); self.analyze(); self.confirm()
        return QDialog.Accepted
    monkeypatch.setattr(MaterialListDialog, 'exec', accept)
    window.import_material_list()
    assert len(window.items) == 2 and window.documents.count() == 0
    assert window.items[1].requester == 'Revisor' and not window.items[1].approved
    window.toggle_all_items(); assert len(window.selected_rows()) == 2
    monkeypatch.setattr(MaterialListDialog, 'exec', lambda self: QDialog.Rejected)
    window.import_material_list(); assert len(window.items) == 2
    window.close(); dialog.close()



def test_nineteen_imported_select_all_and_approve_in_ui(monkeypatch):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    from PySide6.QtWidgets import QApplication
    from cadastro_pcm.interface import main_window
    from cadastro_pcm.interface.material_list_dialog import MaterialListDialog
    from cadastro_pcm.interface.approval_dialog import ApprovalDialog
    app = QApplication.instance() or QApplication([])
    monkeypatch.setattr(main_window, 'load_settings', lambda: Settings(requester='Revisor'))
    window = main_window.Window()
    def accept(self):
        self.text.setPlainText('\n'.join(f'Produto {n} | 84818095 | {n},00' for n in range(1, 20)))
        self.analyze(); assert self.table.rowCount() == 19
        self.confirm(); return 1
    monkeypatch.setattr(MaterialListDialog, 'exec', accept)
    window.import_material_list(); window.toggle_all_items()
    assert len(window.selected_rows()) == 19 and not any(i.approved for i in window.items)
    monkeypatch.setattr(ApprovalDialog, 'exec', lambda self: ApprovalDialog.COMPLETE_ONLY)
    window.approve_selected()
    assert len(window.items) == 19 and all(i.approved for i in window.items)
    window.close()



def test_tabs_missing_description_and_invalid_fields():
    items, notices = parse_material_list('Cadastro\tNCM\tPreço unitário\n\t84818095\t120,00\nProduto\t123\tinválido')
    assert len(items) == 1 and notices
    assert items[0].name == 'Produto' and not items[0].ncm and items[0].unit_price is None
    assert items[0].evidence['ncm'].excerpt == '123'
    assert items[0].warnings
