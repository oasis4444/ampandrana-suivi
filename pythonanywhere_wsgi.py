# Copiez ce contenu dans le fichier WSGI de PythonAnywhere
# (onglet « Web » > « WSGI configuration file »), en remplaçant VOTRE_NOM.
import os
import sys

project = "/home/VOTRE_NOM/ampandrana-suivi"
if project not in sys.path:
    sys.path.insert(0, project)

# Optionnel : une clé secrète fixe (sinon elle est créée dans instance/secret_key)
# os.environ["SECRET_KEY"] = "une-longue-phrase-secrete"

from app import app as application  # noqa: E402,F401
