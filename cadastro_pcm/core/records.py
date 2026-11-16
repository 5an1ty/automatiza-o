"""Central material records, with explicit read compatibility for legacy outputs."""
import hashlib
import json
import os
from pathlib import Path
from uuid import uuid4


def relative_path(path, root):
    try:
        return os.path.relpath(Path(path).resolve(), Path(root).resolve())
    except ValueError:  # Different Windows drives.
        return str(Path(path).resolve())


def record_key(folder, root):
    identity = os.path.normcase(relative_path(folder, root))
    return hashlib.sha256(identity.encode('utf-8')).hexdigest()[:24]


def matches(data, folder, root):
    expected = (Path(folder).resolve() / 'cadastro.xlsx')
    paths = [data.get('excel_path', '')]
    if data.get('excel_relative_path'):
        paths.append(str(root / data['excel_relative_path']))
    return any(path and os.path.normcase(str(Path(path).resolve())) == os.path.normcase(str(expected)) for path in paths)


def material_record(folder, output=None):
    folder = Path(folder).resolve()
    roots = ([Path(output).resolve()] if output is not None else []) + [folder] + list(folder.parents)
    for root in roots:
        directory = root / 'registros json'
        candidate = directory / f'registro_{record_key(folder, root)}.json'
        if candidate.is_file() and matches(json.loads(candidate.read_text(encoding='utf-8')), folder, root):
            return candidate
    # Read only: previous versions placed the record beside the workbook.
    legacy = folder / 'registro.json'
    if legacy.is_file():
        return legacy
    raise FileNotFoundError('Registro do material não encontrado em registros json nem na saída antiga.')


def read_record(folder, output=None):
    return json.loads(material_record(folder, output).read_text(encoding='utf-8'))


def update_record(path, data):
    path = Path(path)
    temporary = path.with_name('.registro_' + uuid4().hex + '.tmp')
    handle = temporary.open('x', encoding='utf-8')
    try:
        with handle:
            handle.write(json.dumps(data, ensure_ascii=False, indent=2))
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def ensure_record(folder, output, data=None):
    folder, output = Path(folder).resolve(), Path(output).resolve()
    directory = output / 'registros json'; directory.mkdir(parents=True, exist_ok=True)
    target = directory / f'registro_{record_key(folder, output)}.json'
    if target.exists():
        existing = json.loads(target.read_text(encoding='utf-8'))
        if not matches(existing, folder, output):
            raise FileExistsError('Identificador de registro já utilizado por outro cadastro.')
        return target
    data = dict(data if data is not None else read_record(folder, output))
    data.update(record_id=record_key(folder, output), excel_path=str(folder / 'cadastro.xlsx'),
                excel_relative_path=relative_path(folder / 'cadastro.xlsx', output))
    temporary = directory / ('.registro_' + uuid4().hex + '.tmp')
    handle = temporary.open('x', encoding='utf-8')
    try:
        with handle:
            handle.write(json.dumps(data, ensure_ascii=False, indent=2))
        if os.name == 'nt':
            os.rename(temporary, target)
        else:
            os.link(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)
    return target
