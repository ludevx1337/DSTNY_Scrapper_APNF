from __future__ import annotations

import logging
import os
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Final

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


LOGGER = logging.getLogger("societe_apnf")

PHONE_COLUMN: Final = "numéro"
SEARCH_FIELD_ID: Final = "tx_numberSearch"
LOGIN_FIELD_ID: Final = "login"
PASSWORD_FIELD_ID: Final = "password"
WAITING_INDICATOR_ID: Final = "waiting-indicator"

RESULT_LABELS: Final[dict[str, str]] = {
    "type": "Type de ligne",
    "derniere_action": "Dernière action",
    "operateur_attributaire": "Opérateur Attributaire",
    "operateur_exploitant": "Opérateur Exploitant",
    "operateur_telecom": "Opérateur Commercial",
}

DATE_RE: Final = re.compile(
    r"\b(?:"
    r"(?P<dmy>\d{1,2}[/-]\d{1,2}[/-]\d{2,4})"
    r"|(?P<ymd>\d{4}-\d{1,2}-\d{1,2})"
    r")\b"
)

PORTABILITY_RE: Final = re.compile(r"\bport(?:a(?:bilit[ée])?)?\b", re.IGNORECASE)

OPERATOR_FLOW_PATTERNS: Final[tuple[re.Pattern[str], ...]] = (
    re.compile(
        r"(?:de|depuis)\s+(?P<from>[^|;\n]+?)\s+(?:vers|à|->|→)\s+"
        r"(?P<to>[^|;\n]+?)(?=\s+(?:le|du)\s+\d|[|;\n]|$)",
        re.IGNORECASE,
    ),
    re.compile(
        r"(?P<from>[^|;\n]+?)\s*(?:->|→)\s*(?P<to>[^|;\n]+?)(?=[|;\n]|$)",
        re.IGNORECASE,
    ),
)


@dataclass(slots=True)
class QualificationResult:
    type: str = ""
    derniere_action: str = ""
    operateur_attributaire: str = ""
    operateur_exploitant: str = ""
    operateur_telecom: str = ""
    date_portabilite_recente: str = ""
    operateur_origine: str = ""
    operateur_destination: str = ""
    statut_recherche: str = "OK"
    erreur_recherche: str = ""


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "oui", "on"}


def get_qualification_url() -> str:
    url = os.getenv("SOCIETE_QUALIFICATION_URL", "").strip()
    if not url:
        raise RuntimeError(
            "La variable SOCIETE_QUALIFICATION_URL est obligatoire. "
            "Consultez le README.md pour la configuration."
        )
    return url


def choose_excel_file() -> Path | None:
    """Ouvre le sélecteur de fichier natif sans dépendance GUI externe."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title="Sélectionner le fichier Excel de leads",
        filetypes=[("Fichiers Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")],
    )
    root.destroy()
    return Path(selected) if selected else None


def normalize_phone(value: object) -> str | None:
    """Normalise un numéro français sur 10 chiffres en conservant le 0 initial."""
    if pd.isna(value):
        return None

    raw = str(value).strip()
    if raw.endswith(".0") and raw[:-2].isdigit():
        raw = raw[:-2]

    digits = re.sub(r"\D+", "", raw)

    if digits.startswith("0033"):
        digits = "0" + digits[4:]
    elif digits.startswith("33") and len(digits) == 11:
        digits = "0" + digits[2:]

    if len(digits) == 9:
        digits = "0" + digits

    if len(digits) == 10 and digits.startswith("0"):
        return digits

    return None


def make_driver() -> webdriver.Remote:
    browser = os.getenv("SOCIETE_BROWSER", "firefox").strip().lower()
    headless = env_bool("SOCIETE_HEADLESS")

    if browser == "chrome":
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("--headless=new")
        return webdriver.Chrome(options=options)

    if browser != "firefox":
        raise RuntimeError("SOCIETE_BROWSER doit valoir 'firefox' ou 'chrome'.")

    options = webdriver.FirefoxOptions()
    if headless:
        options.add_argument("-headless")
    return webdriver.Firefox(options=options)


def wait_for_search_field(driver: webdriver.Remote, timeout: int = 60):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, SEARCH_FIELD_ID))
    )


def login_if_needed(driver: webdriver.Remote, url: str) -> None:
    driver.get(url)

    try:
        login_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, LOGIN_FIELD_ID))
        )
    except TimeoutException:
        wait_for_search_field(driver)
        return

    password_field = driver.find_element(By.ID, PASSWORD_FIELD_ID)
    username = os.getenv("SOCIETE_LOGIN", "").strip()
    password = os.getenv("SOCIETE_PASSWORD", "")

    if username and password:
        login_field.clear()
        login_field.send_keys(username)
        password_field.clear()
        password_field.send_keys(password, Keys.RETURN)
        wait_for_search_field(driver)
        LOGGER.info("Connexion automatisée réussie.")
        return

    LOGGER.warning(
        "SOCIETE_LOGIN/SOCIETE_PASSWORD absents : connectez-vous manuellement "
        "dans le navigateur. Le script reprendra dès que la zone de recherche sera disponible."
    )
    wait_for_search_field(driver, timeout=180)


def xpath_literal(text: str) -> str:
    """Échappe une chaîne pour l'utiliser comme littéral XPath."""
    if "'" not in text:
        return f"'{text}'"
    if '"' not in text:
        return f'"{text}"'
    parts = text.split("'")
    return "concat(" + ', "\'", '.join(f"'{part}'" for part in parts) + ")"


def strip_label(text: str, label: str) -> str:
    cleaned = text.strip()
    pattern = re.compile(rf"^\s*{re.escape(label)}\s*[:\-]?\s*", re.IGNORECASE)
    return pattern.sub("", cleaned, count=1).strip()


def read_labeled_value(
    driver: webdriver.Remote,
    label: str,
    timeout: int = 30,
) -> str:
    label_xpath = xpath_literal(label)
    locator = (
        By.XPATH,
        f"//p[contains(normalize-space(.), {label_xpath})]",
    )
    element = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located(locator)
    )
    return strip_label(element.text, label)


def parse_portability(last_action: str) -> tuple[str, str, str]:
    """Extrait, sans l'inventer, date et flux opérateur d'une action de portabilité."""
    if not last_action or not PORTABILITY_RE.search(last_action):
        return "", "", ""

    date_match = DATE_RE.search(last_action)
    date_value = date_match.group(0) if date_match else ""

    for pattern in OPERATOR_FLOW_PATTERNS:
        match = pattern.search(last_action)
        if match:
            return (
                date_value,
                match.group("from").strip(" :-"),
                match.group("to").strip(" :-"),
            )

    return date_value, "", ""


def dismiss_waiting_indicator(driver: webdriver.Remote) -> None:
    """Attend la disparition du loader et applique un fallback compatible avec l'ancien portail."""
    try:
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.ID, WAITING_INDICATOR_ID))
        )
    except TimeoutException:
        LOGGER.debug("Loader toujours visible ; application du fallback JavaScript.")
        driver.execute_script(
            "const el = document.getElementById(arguments[0]); "
            "if (el) { el.style.display = 'none'; }",
            WAITING_INDICATOR_ID,
        )


def qualify_number(driver: webdriver.Remote, number: str) -> QualificationResult:
    search = wait_for_search_field(driver)
    search.send_keys(Keys.CONTROL, "a")
    search.send_keys(Keys.DELETE)
    search.send_keys(number, Keys.RETURN)

    dismiss_waiting_indicator(driver)

    values = {
        field: read_labeled_value(driver, label)
        for field, label in RESULT_LABELS.items()
    }
    port_date, origin, destination = parse_portability(values["derniere_action"])

    return QualificationResult(
        **values,
        date_portabilite_recente=port_date,
        operateur_origine=origin,
        operateur_destination=destination,
    )


def build_output_path(source: Path) -> Path:
    return source.with_name(f"{source.stem}_apnf.xlsx")


def process_file(driver: webdriver.Remote, source: Path, url: str) -> Path:
    data = pd.read_excel(source, engine="openpyxl")

    if PHONE_COLUMN not in data.columns:
        columns = ", ".join(map(str, data.columns))
        raise KeyError(
            f"Colonne obligatoire '{PHONE_COLUMN}' introuvable. "
            f"Colonnes détectées : {columns or '(aucune)'}."
        )

    results: list[dict[str, str]] = []
    total = len(data)

    for index, raw_number in enumerate(data[PHONE_COLUMN], start=1):
        number = normalize_phone(raw_number)
        LOGGER.info("[%s/%s] Traitement : %s", index, total, number or raw_number)

        if number is None:
            results.append(
                asdict(
                    QualificationResult(
                        statut_recherche="NUMERO_INVALIDE",
                        erreur_recherche="Numéro absent ou format non reconnu.",
                    )
                )
            )
            continue

        try:
            result = qualify_number(driver, number)
            results.append(asdict(result))
        except Exception as exc:  # Une ligne en erreur ne doit pas perdre tout le fichier.
            LOGGER.exception("Échec de la qualification du numéro %s", number)
            results.append(
                asdict(
                    QualificationResult(
                        statut_recherche="ERREUR",
                        erreur_recherche=f"{type(exc).__name__}: {exc}",
                    )
                )
            )
        finally:
            # Le portail historique est plus fiable après rechargement entre deux recherches.
            driver.get(url)

    enriched = pd.concat(
        [data.reset_index(drop=True), pd.DataFrame(results)],
        axis=1,
    )

    output = build_output_path(source)
    enriched.to_excel(output, index=False, engine="openpyxl")
    return output


def main() -> int:
    configure_logging()

    try:
        url = get_qualification_url()
        source = choose_excel_file()
        if source is None:
            LOGGER.info("Aucun fichier sélectionné. Arrêt.")
            return 0

        LOGGER.info("Fichier sélectionné : %s", source)
        LOGGER.info("Démarrage du navigateur.")
        driver = make_driver()

        try:
            login_if_needed(driver, url)
            output = process_file(driver, source, url)
        finally:
            driver.quit()

        LOGGER.info("Opération terminée : %s", output)
        print(f"\nFichier généré : {output}")
        return 0

    except KeyboardInterrupt:
        LOGGER.warning("Opération interrompue par l'utilisateur.")
        return 130
    except Exception as exc:
        LOGGER.exception("Erreur fatale : %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
