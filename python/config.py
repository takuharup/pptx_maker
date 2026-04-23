"""Settings loader — yaml preferred, falls back to built-in defaults."""
import os

_DEFAULTS = {
    'paths': {
        'db_path': 'data/work/database.db',
        'python_exe': 'python',
        'input_dir': 'data/input',
        'work_dir': 'data/work',
        'output_dir': 'data/output',
    }
}


def _load_yaml(path: str) -> dict:
    import yaml
    with open(path, encoding='utf-8') as f:
        return yaml.safe_load(f)


def _load_json(path: str) -> dict:
    import json
    with open(path, encoding='utf-8') as f:
        return json.load(f)


class Config:
    def __init__(self, config_path: str = 'config/settings.yaml') -> None:
        if os.path.isfile(config_path):
            try:
                self._data = _load_yaml(config_path)
            except ImportError:
                json_path = config_path.replace('.yaml', '.json')
                if os.path.isfile(json_path):
                    self._data = _load_json(json_path)
                else:
                    self._data = _DEFAULTS
        else:
            self._data = _DEFAULTS

    def _p(self, key: str) -> str:
        return self._data.get('paths', {}).get(key, _DEFAULTS['paths'][key])

    @property
    def db_path(self) -> str:
        return self._p('db_path')

    @property
    def python_exe(self) -> str:
        return self._p('python_exe')

    @property
    def work_dir(self) -> str:
        return self._p('work_dir')

    @property
    def output_dir(self) -> str:
        return self._p('output_dir')
