# OpenInDocs

**Double-cliquez sur un fichier Word, Excel ou PowerPoint sous Windows et modifiez-le dans Google Docs, Sheets ou Slides.**

[English](README.md)

Google Drive pour ordinateur monte votre Drive comme un disque, mais ne permet pas d'ouvrir un `.docx` local dans l'éditeur Google :
Windows ne propose que Word. OpenInDocs comble ce manque, **sans API Google, sans clé d'API et sans connexion supplémentaire**.

> Projet non affilié à Google et non approuvé par Google. « Google Docs », « Sheets », « Slides » et « Drive » sont des marques de Google LLC.

## Ce que ça fait

1. Vous double-cliquez sur un fichier (par exemple un `contrat.docx` tout juste téléchargé).
2. Une petite fenêtre indique l'avancement de l'envoi. Le fichier est copié dans un dossier de votre Drive (`Mon Drive\Ouverts depuis le PC` par défaut) ; l'original reste où il est.
3. Dès que Google Drive pour ordinateur l'a envoyé, le fichier s'ouvre dans **Google Docs** (`.docx`), **Sheets** (`.xlsx`) ou **Slides** (`.pptx`) dans votre navigateur par défaut.

Un fichier déjà dans votre Drive s'ouvre tout de suite, sans copie. Rouvrir le même fichier ne crée jamais de doublon.
En cas de problème (hors connexion, Drive en pause...), l'outil propose d'ouvrir le fichier dans Word / Excel / PowerPoint.

## Prérequis

- Windows 10 ou 11
- [Google Drive pour ordinateur](https://www.google.com/drive/download/) installé et connecté
- Rien d'autre : l'installateur utilise le compilateur C# et la bibliothèque SQLite fournis avec Windows

## Installation

1. Téléchargez `OpenInDocs.zip` depuis la [dernière release](../../releases/latest) et décompressez-le (ou clonez le dépôt).
2. Clic droit sur `installer\Install.ps1` > **Exécuter avec PowerShell**. Pas besoin d'être administrateur.
3. La page *Applications par défaut* de Windows s'ouvre : choisissez **Google Docs**, **Google Sheets** et **Google Slides** comme applications par défaut de leurs types de fichiers.
   (Autre méthode : double-clic sur un `.docx`, choisir *Google Docs*, puis *Toujours*.)

Pour revenir à Word pour un seul fichier : clic droit > *Ouvrir avec* > *Word*.
Pour désinstaller : `powershell -ExecutionPolicy Bypass -File installer\Install.ps1 -Uninstall`.

Lancez `%LOCALAPPDATA%\OpenInDocs\OpenInDocs.exe` sans fichier pour obtenir un diagnostic de votre installation.

## Réglages

`%LOCALAPPDATA%\OpenInDocs\OpenInDocs.ini` (voir [`config/OpenInDocs.ini`](config/OpenInDocs.ini)) :

| Réglage | Signification |
|---|---|
| `TargetFolder` | Dossier de Mon Drive où les fichiers sont copiés |
| `GoogleAccount` | Compte à utiliser si le navigateur est connecté à plusieurs comptes Google |
| `MaxWaitSeconds` | Attente maximale de l'envoi |
| `DriveRoot` | Emplacement de Mon Drive, si la détection automatique échoue |

## Icônes

Les fichiers prennent les icônes Docs / Sheets / Slides. Elles ne sont **pas** fournies dans ce dépôt : l'installateur les extrait
de votre propre installation de Google Drive pour ordinateur. Si elles sont introuvables ou non vérifiables, des icônes par défaut sont utilisées
et tout le reste fonctionne pareil.

## Fonctionnement

Voir [docs/HOW_IT_WORKS.md](docs/HOW_IT_WORKS.md). En bref : Drive pour ordinateur tient une base SQLite locale de vos fichiers ;
OpenInDocs la lit (en lecture seule) pour retrouver l'identifiant Drive du fichier copié, puis ouvre `docs.google.com/document/d/<ID>/edit`.

## Limites

- **Repose sur une base non documentée.** Une mise à jour de Drive pour ordinateur peut changer son format et casser la recherche. L'outil bascule alors sur Office au lieu d'échouer en silence. Merci d'ouvrir une issue le cas échéant.
- Développé et testé sous Windows 11 en français, avec un seul compte Google et Drive en mode *Diffuser des fichiers*. Le mode *Mettre en miroir* et les autres langues devraient fonctionner mais n'ont pas été testés. Les retours sont les bienvenus.
- Les `.docx` s'ouvrent dans le mode de compatibilité Office de Google et restent des `.docx` sur Drive (*Fichier > Enregistrer au format Google Docs* pour convertir).
- Les fichiers sont envoyés sur **votre** Drive. À éviter pour des fichiers qui ne doivent pas y être stockés.
- Windows 10/11 interdisent aux programmes de définir eux-mêmes les applications par défaut : un choix manuel par type de fichier reste nécessaire.

## Confidentialité

OpenInDocs n'envoie rien lui-même : ni télémétrie, ni accès réseau. Le fichier est envoyé par Google Drive pour ordinateur,
et votre navigateur ouvre une adresse Google. La base locale est ouverte en lecture seule uniquement.

## Contribuer

Issues et pull requests bienvenues, notamment pour d'autres langues de Windows, d'autres versions de Drive et d'autres types de comptes
(Workspace, drives partagés). L'intégration continue compile les trois lanceurs sous Windows à chaque modification.

## Licence

[MIT](LICENSE)
