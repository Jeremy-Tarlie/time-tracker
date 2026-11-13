# Time Tracker

Application de suivi du temps par projet. Interface compacte, system tray, export Excel.

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## Fonctionnalités

- Chronomètre **Start / Pause / Stop** par projet
- Gestion des projets (ajouter, renommer, supprimer)
- Tri des projets (alphabétique ou par date)
- Export Excel (`heures_projet.xlsx`) — feuilles **Sessions** et **Résumé par projet**
- File d'attente locale si Excel est verrouillé (réessai)
- Reprise d'une session après crash / redémarrage
- Historique du jour et saisie manuelle
- System tray (réduire via la croix, quitter via le menu tray)
- Chemin Excel personnalisable
- Notes optionnelles à la fin de chaque session
- **Windows** : installation AppData, raccourcis et démarrage Windows en **opt-in**
- Mode **portable** (`portable.flag` ou `--portable`)
- **Mises à jour automatiques** (binaires uniquement) via GitHub Releases

## Téléchargement

Les binaires sont publiés automatiquement sur la page [**Releases**](https://github.com/Jeremy-Tarlie/time-tracker/releases) à chaque nouvelle version.

| Fichier | Plateforme |
|---------|------------|
| `TimeTracker-windows-x64.exe` | Windows 64-bit |
| `TimeTracker-linux-x64` | Linux 64-bit |
| `TimeTracker-macos-arm64` | macOS Apple Silicon (M1/M2/M3…) |

> macOS Intel : non fourni (runners CI trop lents). Utilisez les sources ou un Mac Apple Silicon.

### Windows

1. Téléchargez `TimeTracker-windows-x64.exe`
2. Lancez-le une fois : l’app propose raccourcis / démarrage Windows (décoché par défaut)
3. Utilisez ensuite le raccourci **Time Tracker** (Bureau / Menu Démarrer)

Mode portable : placez `portable.flag` à côté de l’exe, ou lancez avec `--portable` (pas de copie AppData / pas de registre).

### Linux / macOS

```bash
chmod +x TimeTracker-linux-x64   # ou TimeTracker-macos-arm64
./TimeTracker-linux-x64
```

Données utilisateur :

| OS | Dossier |
|----|---------|
| Windows | `%LOCALAPPDATA%\TimeTracker\` |
| macOS | `~/Library/Application Support/TimeTracker/` |
| Linux | `~/.local/share/TimeTracker/` |
| Portable | dossier de l’exécutable |

Fichiers utiles : `projects.json`, `pending_sessions.json`, `active_session.json`, `timetracker.log`.

## Développement

**Prérequis :** Python 3.10+ (tkinter inclus ; sous Linux : `python3-tk`).

```bash
python -m venv .venv
# Windows
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

En mode développement, les données (`heures_projet.xlsx`, `projects.json`) sont créées à côté de `main.py`.

### Tests / qualité

```bash
.venv\Scripts\python.exe -m unittest discover -s tests -v
.venv\Scripts\python.exe -m ruff check .
.venv\Scripts\python.exe -m mypy time_tracker
```

### Compilation locale (Windows)

```bash
.venv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller
.venv\Scripts\pyinstaller.exe --onefile --windowed --name TimeTracker --clean --icon assets/icon.ico --hidden-import=pystray --hidden-import=PIL --add-data "assets;assets" main.py
```

L’exécutable est généré dans `dist/`.

## Mises à jour automatiques

Les binaires publiés sur [Releases](https://github.com/Jeremy-Tarlie/time-tracker/releases) se mettent à jour seuls :

- **Au démarrage** : si une version plus récente existe, elle est téléchargée et l’app redémarre (sans confirmation).
- **Pendant l’utilisation** : un bouton **Mise à jour** apparaît ; si un chrono ou une écriture Excel est en cours, une confirmation propose Maintenant / Plus tard.

Le mode développement (`python main.py`) ne vérifie pas les mises à jour.

> macOS : le binaire n’est pas signé ; Gatekeeper peut demander une autorisation manuelle après une mise à jour.

## Publier une nouvelle version

1. Incrémentez `APP_VERSION` dans `time_tracker/constants.py` (ex. `"1.2.1"`)
2. Committez et poussez sur `main`

GitHub Actions lance lint/tests, compile pour Windows / Linux / macOS ARM, et crée la release.

## Désinstallation (Windows)

1. Supprimer `%LOCALAPPDATA%\TimeTracker\`
2. Supprimer les raccourcis **Time Tracker** (Bureau et Menu Démarrer)
3. Supprimer la clé de registre `TimeTracker` dans  
   `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`  
   (ou décochez « Lancer au démarrage » dans Réglages avant)

## Structure

```
time-tracker/
├── main.py
├── assets/                         # icônes tray / exe
├── time_tracker/
│   ├── constants.py
│   ├── paths.py
│   ├── install.py
│   ├── config.py
│   ├── session_store.py
│   ├── instance.py
│   ├── excel.py
│   ├── tray.py
│   ├── widgets.py
│   ├── dialogs.py
│   ├── logging_setup.py
│   └── app.py
├── tests/
├── requirements.txt
├── pyproject.toml
├── LICENSE
├── README.md
└── .github/workflows/release.yml
```

## Dépendances

- [openpyxl](https://openpyxl.readthedocs.io/) — Excel
- [pystray](https://github.com/moses-palmer/pystray) + [Pillow](https://pillow.readthedocs.io/) — system tray
- tkinter — interface (fourni avec Python)
- [PyInstaller](https://pyinstaller.org/) — packaging (CI / build local)

## Licence

MIT — voir [LICENSE](LICENSE).
