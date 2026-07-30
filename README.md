# Time Tracker

Application de suivi du temps par projet. Interface compacte, toujours au premier plan, avec export Excel.

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## Fonctionnalités

- Chronomètre **Start / Stop** par projet
- Gestion des projets (ajouter, renommer, supprimer)
- Tri des projets (alphabétique ou par date)
- Export Excel (`heures_projet.xlsx`) — feuilles **Sessions** et **Résumé par projet**
- Chemin Excel personnalisable
- Notes optionnelles à la fin de chaque session
- Fenêtre compacte, déplaçable, toujours au premier plan
- **Windows** : installation automatique (raccourcis, démarrage au boot)

## Téléchargement

Les binaires sont publiés automatiquement sur la page [**Releases**](https://github.com/Jeremy-Tarlie/time-tracker/releases) à chaque nouvelle version.

| Fichier | Plateforme |
|---------|------------|
| `TimeTracker-windows-x64.exe` | Windows 64-bit |
| `TimeTracker-linux-x64` | Linux 64-bit |
| `TimeTracker-macos-arm64` | macOS Apple Silicon (M1/M2/M3…) |
| `TimeTracker-macos-x64` | macOS Intel |

### Windows

1. Téléchargez `TimeTracker-windows-x64.exe`
2. Lancez-le une fois : l’app s’installe dans `%LOCALAPPDATA%\TimeTracker\`
3. Utilisez ensuite le raccourci **Time Tracker** (Bureau / Menu Démarrer)

### Linux / macOS

```bash
chmod +x TimeTracker-linux-x64   # ou TimeTracker-macos-arm64 / TimeTracker-macos-x64
./TimeTracker-linux-x64
```

Données utilisateur :

| OS | Dossier |
|----|---------|
| Windows | `%LOCALAPPDATA%\TimeTracker\` |
| macOS | `~/Library/Application Support/TimeTracker/` |
| Linux | `~/.local/share/TimeTracker/` |

## Développement

**Prérequis :** Python 3.10+ (tkinter inclus ; sous Linux : `python3-tk`).

```bash
pip install -r requirements.txt
python main.py
```

En mode développement, les données (`heures_projet.xlsx`, `projects.json`) sont créées à côté de `main.py`.

### Compilation locale (Windows)

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name TimeTracker --clean main.py
```

L’exécutable est généré dans `dist/`.

## Publier une nouvelle version

1. Incrémentez `APP_VERSION` dans `time_tracker/constants.py` (ex. `"1.1.3"`)
2. Committez et poussez sur `main`

```bash
git add time_tracker/constants.py
git commit -m "release: v1.1.3"
git push origin main
```

GitHub Actions détecte la nouvelle version, compile pour Windows / Linux / macOS, et crée la release `v1.1.3` avec les binaires.  
Si la release existe déjà pour cette version, le workflow ne republie rien.

## Désinstallation (Windows)

1. Supprimer `%LOCALAPPDATA%\TimeTracker\`
2. Supprimer les raccourcis **Time Tracker** (Bureau et Menu Démarrer)
3. Supprimer la clé de registre `TimeTracker` dans  
   `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`

## Structure

```
time-tracker/
├── main.py                         # Point d'entrée
├── time_tracker/
│   ├── constants.py                # Version, types, options de tri
│   ├── paths.py                    # Chemins données / install
│   ├── install.py                  # Installation Windows
│   ├── config.py                   # projects.json
│   ├── excel.py                    # Export Excel
│   ├── widgets.py                  # SearchableDropdown
│   ├── dialogs.py                  # Fenêtres modales
│   └── app.py                      # Fenêtre principale
├── requirements.txt
├── LICENSE
├── README.md
└── .github/workflows/release.yml
```

## Dépendances

- [openpyxl](https://openpyxl.readthedocs.io/) — Excel
- tkinter — interface (fourni avec Python)
- [PyInstaller](https://pyinstaller.org/) — packaging (CI / build local)

## Licence

MIT — voir [LICENSE](LICENSE).
