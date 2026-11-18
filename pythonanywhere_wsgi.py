# Ajoutez ces lignes À LA FIN du fichier WSGI de PythonAnywhere
# (onglet « Web » > « WSGI configuration file »).
# Le site existant reste à la racine ; Ampandrana Suivi est servi sous /ampandrana.
# Si Ampandrana est le seul site, remplacez la dernière ligne par :  application = _amp.app
import importlib.util as _ilu
import sys as _sys

_AMP_DIR = "/home/VOTRE_NOM/ampandrana-suivi"
if _AMP_DIR not in _sys.path:
    _sys.path.append(_AMP_DIR)
_spec = _ilu.spec_from_file_location("ampandrana_app", _AMP_DIR + "/app.py")
_amp = _ilu.module_from_spec(_spec)
_sys.modules["ampandrana_app"] = _amp
_spec.loader.exec_module(_amp)

from werkzeug.middleware.dispatcher import DispatcherMiddleware  # noqa: E402
application = DispatcherMiddleware(application, {"/ampandrana": _amp.app})  # noqa: F821
