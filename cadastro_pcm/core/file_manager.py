import re
from pathlib import Path
from uuid import uuid4


def safe_name(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')[:100]
    if not name:
        name = 'material'
    if name.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(1,10)), *(f'LPT{i}' for i in range(1,10))}:
        name = '_' + name
    return name


def new_output(root, name):
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    for _ in range(10):
        target = root / (safe_name(name) + '_' + uuid4().hex[:10])
        try:
            target.mkdir()
            return target
        except FileExistsError:
            continue
    raise ValueError('Não foi possível criar uma pasta exclusiva.')
