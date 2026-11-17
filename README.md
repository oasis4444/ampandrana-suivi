# Ampandrana Suivi

Logiciel de suivi des commandes et de la caisse d'Ampandrana création logiciel (Flask + SQLite).

## Menus

| Menu | Contenu |
|---|---|
| **Tableau de bord** | Encaissé par mode (Espèces, MVola, Orange Money, mobile money, autres) pour aujourd'hui, ce mois, le mois précédent, tout ou des dates précises · vola noraisin'ny mpikambana (encaissé par membre et par mode) · nampiditra (qui a saisi) · graphique par mois · liste « À vérifier » · chiffres des commandes |
| **Commandes** | Recherche, filtres, création, modification ; chaque commande montre montant, payé, reste et ses paiements, avec ajout de paiement |
| **Paiements** | Journal des encaissements avec filtres (période, mode, encaissé par, saisi par, à vérifier), modification, annulation avec motif, export CSV |
| **Mr** | Résumé par mois et détail par jour, comme la feuille « Mr » de l'Excel : nombre de messages, nombre de projets, montant des projets, avances reçues, reste à payer, plus l'encaissé par mode ; deux graphiques |
| **Rendez-vous** | Liste et fiche des rendez-vous |
| **Équipe** | Membres proposés dans « Encaissé par », « Responsable »… |
| **Listes** | Statuts de commande, statuts de livraison, modes de paiement, en ligne / local (modifiables par un administrateur) |
| **Utilisateurs** (admin) | Un compte par personne : chaque paiement enregistre automatiquement le compte connecté (« Saisi par ») |
| **Import / Journal** (admin) | Import du fichier Excel, export complet en Excel, journal de toutes les actions |

Règles de contrôle sur les paiements : mode, « Nandray ny vola » et date obligatoires ; référence obligatoire pour MVola, Orange Money et Airtel Money ; le montant ne se modifie pas après coup ; un paiement ne se supprime pas, il s'annule avec un motif et reste visible.

## Lancer sur son ordinateur

```bash
pip install -r requirements.txt
flask --app app run --debug
```

Ouvrez http://127.0.0.1:5000 : la première page crée le compte administrateur.

## Mise en ligne sur PythonAnywhere

1. Créez un compte sur https://www.pythonanywhere.com (le compte gratuit suffit).
2. Ouvrez une console **Bash** (onglet *Consoles*) et récupérez le code :
   ```bash
   git clone https://github.com/VOTRE_COMPTE_GITHUB/ampandrana-suivi.git
   cd ampandrana-suivi
   pip3.10 install --user -r requirements.txt
   ```
   (Pour un dépôt privé, GitHub demande votre nom d'utilisateur et un *personal access token* comme mot de passe.)
3. Onglet **Web** → *Add a new web app* → *Manual configuration* → **Python 3.10**.
4. Dans la section *Code* de l'onglet Web :
   - *Source code* : `/home/VOTRE_NOM/ampandrana-suivi`
   - Cliquez sur le *WSGI configuration file*, effacez tout et collez le contenu de `pythonanywhere_wsgi.py` en remplaçant `VOTRE_NOM`.
5. Section *Static files* : URL `/static/` → dossier `/home/VOTRE_NOM/ampandrana-suivi/static`.
6. Section *Security* : activez **Force HTTPS**.
7. Cliquez sur **Reload**, puis ouvrez `https://VOTRE_NOM.pythonanywhere.com`.
8. Créez le compte administrateur, puis menu **Import / Journal** → choisissez votre fichier Excel (`Ampandrana_Suivi_des_commandes_v3.xlsx`) → **Importer**.
9. Menu **Utilisateurs** : créez un compte pour chaque membre (Sambatra, Toky, Santatra, Princy…).

### Mettre à jour après une modification du code

```bash
cd ~/ampandrana-suivi && git pull
```
puis **Reload** dans l'onglet Web. Les données (dossier `instance/`) ne sont pas touchées.

### Sauvegarde

La base est le fichier `instance/ampandrana.db`. Téléchargez-le de temps en temps depuis l'onglet *Files*, ou utilisez **Exporter tout en Excel** dans le tableau de bord.
Le compte gratuit PythonAnywhere doit être prolongé une fois tous les 3 mois (bouton dans l'onglet Web).

## Import en ligne de commande

```bash
flask --app app import-excel Ampandrana_Suivi_des_commandes_v3.xlsx
```
Remplace les commandes, paiements, rendez-vous et l'équipe par ceux du fichier.

## Fichiers

- `app.py` — application (routes, calculs, droits)
- `schema.sql` — tables SQLite
- `importer.py` / `exporter.py` — import et export Excel
- `templates/`, `static/style.css` — pages
- `pythonanywhere_wsgi.py` — modèle de fichier WSGI

Les données ne sont jamais envoyées sur GitHub : `instance/`, `*.db` et `*.xlsx` sont exclus par `.gitignore`.
