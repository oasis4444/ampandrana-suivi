PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    nom TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'membre',      -- 'admin' ou 'membre'
    actif INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS equipe (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL UNIQUE,
    role TEXT,
    tel TEXT,
    actif INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS listes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    categorie TEXT NOT NULL,                  -- statut_commande, statut_livraison, mode_paiement, canal
    valeur TEXT NOT NULL,
    ordre INTEGER NOT NULL DEFAULT 0,
    UNIQUE (categorie, valeur)
);

CREATE TABLE IF NOT EXISTS commandes (
    id INTEGER PRIMARY KEY,                   -- N° : CMD-001 = 1
    date TEXT,
    client TEXT NOT NULL,
    tel TEXT,
    situation TEXT,
    produit TEXT,
    qte REAL,
    pu REAL,
    statut TEXT,
    livraison TEXT,
    date_livraison TEXT,
    recu_par TEXT,
    responsable TEXT,
    equipe_msg TEXT,
    canal TEXT,
    remarques TEXT,
    created_by INTEGER REFERENCES users(id),
    created_at TEXT,
    updated_by INTEGER REFERENCES users(id),
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS paiements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    commande_id INTEGER NOT NULL REFERENCES commandes(id),
    date TEXT NOT NULL,
    montant REAL NOT NULL,
    mode TEXT NOT NULL,
    reference TEXT,
    encaisse_par TEXT,
    remarque TEXT,
    importe INTEGER NOT NULL DEFAULT 0,
    saisi_par INTEGER REFERENCES users(id),   -- automatique : utilisateur connecté
    saisi_le TEXT,
    modifie_par INTEGER REFERENCES users(id),
    modifie_le TEXT,
    annule INTEGER NOT NULL DEFAULT 0,
    annule_par INTEGER REFERENCES users(id),
    annule_le TEXT,
    motif_annulation TEXT
);
CREATE INDEX IF NOT EXISTS idx_paiements_cmd ON paiements(commande_id);
CREATE INDEX IF NOT EXISTS idx_paiements_date ON paiements(date);
CREATE INDEX IF NOT EXISTS idx_commandes_date ON commandes(date);

CREATE TABLE IF NOT EXISTS rdv (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    projet TEXT NOT NULL,
    tel TEXT,
    lieu TEXT,
    frais TEXT,
    date TEXT,
    heure TEXT,
    responsable TEXT,
    statut TEXT,
    notes TEXT,
    updated_by INTEGER REFERENCES users(id),
    updated_at TEXT
);

-- Journal : chaque création / modification / annulation, avec l'utilisateur
CREATE TABLE IF NOT EXISTS journal (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    quand TEXT NOT NULL,
    user_id INTEGER REFERENCES users(id),
    action TEXT NOT NULL,
    objet TEXT NOT NULL,
    details TEXT
);
