"""New material output contract, legacy compatibility and partial-failure recovery."""
import json
from pathlib import Path
import pytest
from cadastro_pcm.core import generation, file_manager, pdf_recovery
from cadastro_pcm.core.records import material_record, read_record
from cadastro_pcm.core.file_manager import material_pdf, export_central_pdf
from cadastro_pcm.core.settings import Settings
from cadastro_pcm.core.excel_writer import fingerprint, write_excel
from test_workflow import template, approved
from test_central_pdfs import synthetic_export


def settings_for(template, tmp_path):
    return Settings(requester='Revisor', output=str(tmp_path / 'documentos'), template=str(template),
                    authorized_template_sha256=fingerprint(template))


@pytest.mark.parametrize('count', [1, 4])
def test_new_structure_records_and_both_pdf_locations(template, tmp_path, monkeypatch, count):
    settings = settings_for(template, tmp_path)
    monkeypatch.setattr(generation, 'export_pdf', synthetic_export)
    before = template.read_bytes()
    if count == 1:
        folder, error, _ = generation.generate_documents(approved(), settings)
        assert not error
        folders = [folder]
    else:
        _, results = generation.generate_batch([(i, approved()) for i in range(count)], settings)
        assert all(r.status == 'sucesso' for r in results)
        folders = [r.folder for r in results]
    output = Path(settings.output)
    assert len(list((output / 'registros json').glob('*.json'))) == count
    assert len(list((output / 'lote pdf').glob('*.pdf'))) == count
    for folder in folders:
        assert (folder / 'cadastro.xlsx').is_file() and (folder / 'cadastro.pdf').is_file()
        assert not list(folder.glob('*.json'))
        record = material_record(folder)
        assert record.parent == output / 'registros json'
        data = read_record(folder)
        assert data['name'] == approved().name and data['approved']
        assert data['excel_path'] == str(folder.resolve() / 'cadastro.xlsx')
        assert Path(data['pdf_path']).parent == folder.resolve()
        assert Path(data['pdf_shared_path']).parent == output.resolve() / 'lote pdf'
        assert Path(data['pdf_path']).read_bytes() == Path(data['pdf_shared_path']).read_bytes()
        assert data['generation_state'] == {'excel': 'gerado', 'pdf': 'aprovado'}
        assert data['record_id'] in record.name and len(data['pdf_exports']) == 1
        assert pdf_recovery.existing_workbook(folder / 'cadastro.xlsx')['sheets']
    assert template.read_bytes() == before


def test_excel_only_record_outside_material_folder(template, tmp_path):
    settings = settings_for(template, tmp_path)
    folder, error, _ = generation.generate_documents(approved(), settings, pdf=False)
    assert not error and not list(folder.glob('*.json')) and not list(folder.glob('*.pdf'))
    assert read_record(folder)['generation_state'] == {'excel': 'gerado', 'pdf': 'não solicitado'}
    assert material_record(folder).parent == Path(settings.output) / 'registros json'


def test_export_failure_saves_failed_record_and_keeps_excel(template, tmp_path):
    settings = settings_for(template, tmp_path)
    def fail(*args): raise RuntimeError('COM simulado indisponível')
    folder, error, diagnostic = generation.generate_documents(approved(), settings, exporter=fail)
    assert error and diagnostic and (folder / 'cadastro.xlsx').is_file()
    data = read_record(folder)
    assert data['generation_state']['excel'] == 'gerado' and data['generation_state']['pdf'] == 'falhou'
    assert not list(folder.glob('*.json')) and not list(folder.glob('*.pdf'))
    assert not list((Path(settings.output) / 'lote pdf').iterdir())


def test_reexport_keeps_history_and_does_not_overwrite_either_pdf(template, tmp_path):
    settings = settings_for(template, tmp_path)
    folder, _, _ = generation.generate_documents(approved(), settings, exporter=synthetic_export)
    data = read_record(folder)
    before = {Path(data[key]): Path(data[key]).read_bytes() for key in ('pdf_path', 'pdf_shared_path', 'excel_path')}
    record_path = material_record(folder)
    export_central_pdf(folder / 'cadastro.xlsx', folder, settings.output, settings.sheet,
                       approved().name, exporter=synthetic_export)
    updated = read_record(folder)
    assert material_record(folder) == record_path
    assert updated['pdf_path'] != data['pdf_path'] and updated['pdf_shared_path'] != data['pdf_shared_path']
    assert len(updated['pdf_exports']) == 2 and updated['pdf_exports'][0] == data['pdf_exports'][0]
    assert all(path.read_bytes() == contents for path, contents in before.items())
    assert Path(updated['pdf_path']).read_bytes() == Path(updated['pdf_shared_path']).read_bytes()


def test_shared_copy_failure_preserves_local_pdf_and_record_then_retry(template, tmp_path, monkeypatch):
    settings = settings_for(template, tmp_path)
    publish = file_manager.publish_pdf
    def fail_shared(source, directory, **kwargs):
        if Path(directory).name == 'lote pdf': raise PermissionError(13, 'Cópia central bloqueada')
        return publish(source, directory, **kwargs)
    monkeypatch.setattr(file_manager, 'publish_pdf', fail_shared)
    folder, error, _ = generation.generate_documents(approved(), settings, exporter=synthetic_export)
    assert error
    data = read_record(folder)
    local = Path(data['pdf_path']); contents = local.read_bytes()
    assert data['generation_state']['pdf'] == 'falhou' and data['pdf_exports'][0]['shared_path'] == ''
    assert local.parent == folder and not list((Path(settings.output) / 'lote pdf').iterdir())
    monkeypatch.setattr(file_manager, 'publish_pdf', publish)
    monkeypatch.setattr(pdf_recovery, 'export_pdf', synthetic_export)
    result = generation.BatchResult(0, approved(), 'erro', error, folder)
    retried = generation.retry_batch_pdfs([result], settings)
    assert retried[0].status == 'sucesso' and local.read_bytes() == contents
    assert read_record(folder)['generation_state']['pdf'] == 'aprovado'


def test_old_record_is_readable_and_reexport_creates_central_record(template, tmp_path):
    folder = tmp_path / 'saida_antiga' / 'material'; folder.mkdir(parents=True)
    item = approved(); source = folder / 'cadastro.xlsx'
    write_excel(template, source, item, 'Criação de Material Nacional', fingerprint(template))
    legacy = folder / 'registro.json'; legacy.write_text(item.model_dump_json())
    before = legacy.read_bytes(); excel = source.read_bytes()
    assert material_record(folder) == legacy.resolve()
    assert pdf_recovery.existing_workbook(source)['sheets']
    export_central_pdf(source, folder, folder.parent, 'Criação de Material Nacional', item.name, exporter=synthetic_export)
    assert material_record(folder).parent == folder.parent / 'registros json'
    assert legacy.read_bytes() == before and source.read_bytes() == excel
    assert read_record(folder)['approved'] and material_pdf(folder).parent == folder


def test_external_workbook_record_resolved_from_configured_output(template, tmp_path):
    settings = settings_for(template, tmp_path)
    folder = tmp_path / 'outro_local' / 'material'; folder.mkdir(parents=True)
    _, error, _ = generation.generate_documents(approved(), settings, folder=folder, exporter=synthetic_export)
    assert not error and not list(folder.glob('*.json'))
    record = material_record(folder, settings.output)
    assert record.parent == Path(settings.output) / 'registros json'
    local = material_pdf(folder, output=settings.output)
    assert local.is_file()
    result = generation.BatchResult(0, approved(), 'sucesso', folder=folder, record_root=settings.output)
    assert result.details['destination'] == str(local)
    assert pdf_recovery.existing_workbook(folder / 'cadastro.xlsx', record_root=settings.output)['sheets']


def test_same_material_folder_name_in_distinct_batches_has_unique_records(template, tmp_path):
    settings = settings_for(template, tmp_path)
    folders = [Path(settings.output) / f'lote_{n}' / 'material' for n in range(2)]
    for folder in folders:
        folder.mkdir(parents=True)
        _, error, _ = generation.generate_documents(approved(), settings, folder=folder, exporter=synthetic_export)
        assert not error
    assert material_record(folders[0]) != material_record(folders[1])
    assert read_record(folders[0])['excel_path'] != read_record(folders[1])['excel_path']


def test_moved_output_tree_keeps_record_discoverable(template, tmp_path):
    settings = settings_for(template, tmp_path)
    folder, _, _ = generation.generate_documents(approved(), settings, pdf=False)
    new_root = tmp_path / 'documentos_movidos'
    Path(settings.output).rename(new_root)
    moved = new_root / folder.name
    assert material_record(moved).parent == new_root / 'registros json'
    assert pdf_recovery.existing_workbook(moved / 'cadastro.xlsx')['sheets']
