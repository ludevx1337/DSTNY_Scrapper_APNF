# Scraper de qualification de lignes fixes / APNF

Outil Python destiné à enrichir un fichier Excel de leads contenant des numéros de téléphone fixe avec les informations disponibles sur le portail de qualification de la société.

Le script ouvre le fichier Excel, interroge automatiquement le portail pour chaque numéro, récupère les informations de qualification de la ligne et génère un nouveau fichier Excel enrichi.

## Informations récupérées

Pour chaque numéro, l'outil récupère notamment :

- le type de ligne ;
- la dernière action connue ;
- l'opérateur attributaire ;
- l'opérateur exploitant ;
- l'opérateur commercial ;
- la date de portabilité récente lorsqu'elle est explicitement présente dans la dernière action ;
- l'opérateur d'origine et l'opérateur de destination lorsque le texte de la dernière action indique explicitement un flux de portabilité ;
- le statut de la recherche et, en cas d'échec, le détail de l'erreur.

L'objectif est de faciliter l'analyse de données de qualification/APNF à partir d'un tableau de leads, par exemple pour identifier une portabilité récente et, lorsque la donnée source le permet, le passage d'un opérateur à un autre.

> Le script ne déduit pas artificiellement l'opérateur d'origine ou de destination à partir de champs qui n'ont pas ce sens. Ces colonnes restent vides lorsque l'information n'est pas explicitement disponible dans la réponse du portail.

## Fonctionnement

Le fichier source doit contenir au minimum une colonne nommée exactement :

```text
numéro
```

Exemple :

| numéro | client | ville |
|---|---|---|
| 01 23 45 67 89 | Exemple A | Paris |
| 123456789 | Exemple B | Lyon |

Les numéros sont normalisés avant la recherche : espaces et séparateurs sont supprimés, un numéro français sur 9 chiffres récupère automatiquement son `0` initial, et les formats `+33` / `0033` sont convertis en format national.

Le résultat est enregistré à côté du fichier original sous la forme :

```text
nom_du_fichier_apnf.xlsx
```

Le fichier d'origine n'est pas modifié.

## Prérequis

- Python 3.11 ou supérieur ;
- Firefox ou Chrome ;
- un accès autorisé au portail de qualification de la société ;
- un fichier Excel `.xlsx` ou `.xlsm`.

Selenium 4 utilise Selenium Manager pour gérer automatiquement le pilote du navigateur dans les configurations standards. Il n'est donc généralement plus nécessaire de télécharger manuellement `geckodriver` ou `chromedriver`.

## Installation

Sous PowerShell :

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Versions de dépendances prévues par ce dépôt :

```text
selenium 4.49.0
pandas 3.0.5
openpyxl 3.1.5
```

## Configuration

Aucune URL, aucun identifiant et aucun mot de passe ne sont enregistrés dans le code.

Définissez au minimum l'URL du portail dans l'environnement :

```powershell
$env:SOCIETE_QUALIFICATION_URL="https://votre-portail.example/Qualification"
```

Pour automatiser également la connexion :

```powershell
$env:SOCIETE_LOGIN="mon_identifiant"
$env:SOCIETE_PASSWORD="mon_mot_de_passe"
```

Si les variables d'identification ne sont pas définies, le navigateur s'ouvre et laisse jusqu'à trois minutes pour effectuer la connexion manuellement. Le traitement reprend automatiquement dès que la zone de recherche est disponible.

### Choix du navigateur

Firefox est utilisé par défaut :

```powershell
$env:SOCIETE_BROWSER="firefox"
```

Pour Chrome :

```powershell
$env:SOCIETE_BROWSER="chrome"
```

Le mode sans interface graphique est optionnel :

```powershell
$env:SOCIETE_HEADLESS="true"
```

Pour une première utilisation, il est conseillé de conserver le navigateur visible afin de vérifier le comportement du portail.

## Lancement

```powershell
python main.py
```

Une fenêtre permet de sélectionner le fichier Excel à traiter.

Le script affiche ensuite la progression dans la console et génère le fichier enrichi à côté du fichier source.

## Colonnes ajoutées

| Colonne | Description |
|---|---|
| `type` | Type de ligne retourné par le portail |
| `derniere_action` | Dernière action connue |
| `operateur_attributaire` | Opérateur attributaire |
| `operateur_exploitant` | Opérateur exploitant |
| `operateur_telecom` | Opérateur commercial |
| `date_portabilite_recente` | Date détectée dans une action explicitement liée à une portabilité |
| `operateur_origine` | Opérateur source si le flux est explicitement indiqué |
| `operateur_destination` | Opérateur destination si le flux est explicitement indiqué |
| `statut_recherche` | `OK`, `NUMERO_INVALIDE` ou `ERREUR` |
| `erreur_recherche` | Détail technique si une ligne n'a pas pu être traitée |

## Améliorations par rapport au script historique

- API Selenium moderne avec `By` et `WebDriverWait` ;
- suppression des anciennes méthodes `find_element_by_*` ;
- suppression des attentes implicites et des `sleep` inutiles ;
- suppression des imports et dépendances non utilisés ;
- gestion automatique Firefox / Chrome ;
- configuration externalisée via variables d'environnement ;
- aucune donnée de connexion dans le dépôt ;
- normalisation plus robuste des numéros français ;
- traitement des erreurs ligne par ligne sans perdre tout le fichier ;
- conservation de toutes les lignes et de leur ordre initial ;
- extraction APNF/portabilité prudente, sans inventer une information absente ;
- chemin de sortie propre avec suffixe `_apnf.xlsx` ;
- fermeture garantie du navigateur avec `driver.quit()`.

## Sécurité et utilisation

Cet outil doit uniquement être utilisé sur un portail et des données auxquels vous êtes autorisé à accéder. Ne placez jamais d'identifiants, mots de passe, cookies de session ou fichiers de leads réels dans un dépôt Git public.

Le `.gitignore` fourni exclut notamment les fichiers Excel générés et les fichiers d'environnement locaux.

## Licence

Ce projet est distribué sous licence **MIT**.

La licence MIT est adaptée à un petit outil open source : elle autorise l'utilisation, la modification, la redistribution et l'intégration du code dans d'autres projets, y compris commerciaux, à condition de conserver la notice de licence.

Voir le fichier [`LICENSE`](LICENSE).
