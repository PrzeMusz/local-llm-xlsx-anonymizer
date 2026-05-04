# -*- coding: utf-8 -*-
"""
analyze.py - skanuje .xlsx, identyfikuje encje przez LLM,
generuje entities.xlsx z pseudonimami do edycji.

Uzycie:
    python analyze.py                    (autodetekcja)
    python analyze.py moj_plik.xlsx

Wynik: entities.xlsx
  Kolumny: Kategoria | Wartosc | Pseudonim | Ile_razy | Anonimizuj
  - Pseudonim: automatyczna propozycja (3+3 liter dla osob, 3 liter dla firm)
  - Ile_razy: ile razy encja wystepuje w Excelu (pomaga ocenic waznosc)
  - Edytuj kolumne "Anonimizuj" (TAK/NIE) i "Pseudonim" wg potrzeb

Nastepnie: python anonymize.py moj_plik.xlsx
"""

import os
import re
import sys
import json
import time
import glob
import unicodedata
import requests
from typing import List, Dict, Tuple
from collections import OrderedDict, Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ======================== KONFIG ========================

LM_URL        = "http://localhost:1234/v1/chat/completions"
MODEL         = "google/gemma-4-e2b"
API_KEY       = "lm-studio"
BATCH_SIZE    = 6
WORKERS       = 4
TIMEOUT_S     = 120
MAX_RETRIES   = 2
MIN_LEN       = 3
ENTITIES_FILE = "entities.xlsx"

GROUPS = OrderedDict([
    ("OSOBA",   "Imie i nazwisko"),
    ("FIRMA",   "Firma / marka / projekt / produkt"),
    ("MIEJSCE", "Adres / miasto / ulica / kod pocztowy"),
    ("DANE",    "Telefon / email / NIP / PESEL / REGON / konto"),
])

GROUP_COLORS = {
    "OSOBA":   "4472C4",
    "FIRMA":   "ED7D31",
    "MIEJSCE": "70AD47",
    "DANE":    "7030A0",
}

SKIP_RE  = re.compile(r"[\d\s\.\,\-\+\/:%]+")

# Sufiksy datowe/liczbowe do odciecia
_STRIP_SUFFIXES = [
    re.compile(r"[\s\-_]*20\d{2}\s*[\-\s/]*\d{0,2}\s*$"),
    re.compile(r"(?<=[a-zA-ZąćęłńóśźżĄĆĘŁŃÓŚŹŻ])\d{3,}\s*$"),
    re.compile(r"\s*[\-/]\s*\d{1,2}\s*$"),
]
_MIN_LEN_AFTER_STRIP = 2

_PYTHON_LIST_RE = re.compile(r"^\[['\"](.*?)['\"]\]$")

SESSION = requests.Session()
SESSION.headers.update({"Authorization": f"Bearer {API_KEY}"})


# ==================== NORMALIZACJA ======================

def clean_cell_value(s: str) -> str:
    s = s.strip()
    m = _PYTHON_LIST_RE.match(s)
    if m:
        return m.group(1).strip()
    return s


def normalize_entity(val: str) -> str:
    result = val.strip()
    changed = True
    while changed:
        changed = False
        for pattern in _STRIP_SUFFIXES:
            new = pattern.sub("", result).strip(" ,-./")
            if new != result and len(new) >= _MIN_LEN_AFTER_STRIP:
                result = new
                changed = True
    return result


def strip_accents(s: str) -> str:
    """Zamienia polskie znaki na ASCII: ł->l, ą->a, itd."""
    # Specjalna obsluga ł/Ł bo unicodedata nie konwertuje ich dobrze
    s = s.replace("ł", "l").replace("Ł", "L")
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def make_pseudonym(entity: str, category: str, used: set) -> str:
    """
    Generuje czytelny pseudonim:
    - OSOBA: 3 litery imienia + _ + 3 litery nazwiska -> Prz_Oma
    - FIRMA/inne: 3 litery nazwy -> Ave, Poc, Nex
    - Wieloczlonowe firmy: 3+3 -> Ave_Den
    Jesli kolizja - dodaje cyfre.
    """
    parts = entity.split()
    ascii_parts = [strip_accents(p) for p in parts]

    if category == "OSOBA" and len(ascii_parts) >= 2:
        base = ascii_parts[0][:3] + "_" + ascii_parts[-1][:3]
    elif len(ascii_parts) >= 2 and category in ("FIRMA", "MIEJSCE"):
        base = ascii_parts[0][:3] + "_" + ascii_parts[1][:3]
    else:
        base = strip_accents(entity.replace(" ", ""))[:5]

    # Pierwsza litera wielka, reszta mala
    base = base[0].upper() + base[1:] if base else "X"

    # Unikaj kolizji
    candidate = base
    counter = 2
    while candidate.lower() in used:
        candidate = f"{base}{counter}"
        counter += 1
    used.add(candidate.lower())
    return candidate


# ==================== LLM ==============================

SYSTEM_PROMPT = (
    "Wyodrebniasz nazwy wlasne z tekstu: imiona, nazwiska, nazwy firm, marek, produktow, projektow. "
    "Tekst moze byc sklejony (np. 'Batch2024-04NazwaFirmy') - wyciagaj TYLKO nazwy wlasne. "
    "Slowka techniczne jak Batch, Monthly, Regular, Rework - wyciagaj osobno jesli sa nazwe projektu. "
    "Daty, lata, numery - POMIJAJ. "
    "Zwracasz WYLACZNIE JSON bez markdown."
)

FEW_SHOTS = [
    (
        "Wyodrebnij encje:\n"
        "1) Jan Kowalski z firmy ACME Sp. z o.o. zadzwonil pod 123-456-789.\n"
        "2) Faktura dla Anny Nowak, ul. Slowackiego 12, 38-400 Krosno.",
        '{"OSOBA":["Jan Kowalski","Anna Nowak"],'
        '"FIRMA":["ACME Sp. z o.o."],'
        '"MIEJSCE":["ul. Slowackiego 12, 38-400 Krosno"],'
        '"DANE":["123-456-789"]}',
    ),
    (
        "Wyodrebnij encje:\n"
        "1) Batch2024 - 04Avery Dennison Rework\n"
        "2) Monthly Regular2025 - 10Nextbase\n"
        "3) Monthly Regular2025 - 10Pocketbook Rework\n"
        "4) Zamowienie gotowe do odbioru.",
        '{"OSOBA":[],'
        '"FIRMA":["Avery Dennison","Rework","Nextbase","Pocketbook"],'
        '"MIEJSCE":[],'
        '"DANE":[]}',
    ),
    (
        "Wyodrebnij encje:\n"
        "1) [\'Inhalatory wodoru2023 - 03\']\n"
        "2) Monthly Regular2025 - 10LG\n"
        "3) Kontakt: biuro@firma.pl, NIP 1234567890.",
        '{"OSOBA":[],'
        '"FIRMA":["Inhalatory wodoru","LG"],'
        '"MIEJSCE":[],'
        '"DANE":["biuro@firma.pl","NIP 1234567890"]}',
    ),
]


def _base_messages() -> List[dict]:
    msgs = [{"role": "system", "content": SYSTEM_PROMPT}]
    for u, a in FEW_SHOTS:
        msgs.append({"role": "user",      "content": u})
        msgs.append({"role": "assistant", "content": a})
    return msgs

BASE_MESSAGES = _base_messages()


def build_messages(texts: List[str]) -> List[dict]:
    msgs = list(BASE_MESSAGES)
    lines = "\n".join(f"{i+1}) {t}" for i, t in enumerate(texts))
    msgs.append({"role": "user", "content": f"Wyodrebnij encje:\n{lines}"})
    msgs.append({"role": "assistant", "content": '{"OSOBA":'})
    return msgs


def parse_entities(raw: str) -> Dict[str, List[str]]:
    raw = raw.strip()
    if not raw.startswith("{"):
        raw = '{"OSOBA":' + raw
    raw = re.sub(r"```[a-zA-Z]*\n?", "", raw).replace("```", "").strip()
    start = raw.find("{")
    if start < 0:
        return {}
    depth, end = 0, start
    for i in range(start, len(raw)):
        if raw[i] == "{": depth += 1
        elif raw[i] == "}":
            depth -= 1
            if depth == 0:
                end = i + 1
                break
    try:
        data = json.loads(raw[start:end])
    except json.JSONDecodeError:
        chunk = re.sub(r",\s*}", "}", raw[start:end])
        try:
            data = json.loads(chunk)
        except json.JSONDecodeError:
            return {}

    result = {}
    for key in GROUPS:
        vals = data.get(key, [])
        if isinstance(vals, list):
            result[key] = [str(v).strip() for v in vals if str(v).strip()]
        elif isinstance(vals, str) and vals.strip():
            result[key] = [vals.strip()]
        else:
            result[key] = []
    return result


def call_llm(texts: List[str]) -> Dict[str, List[str]]:
    if not texts:
        return {k: [] for k in GROUPS}
    payload = {
        "model": MODEL,
        "messages": build_messages(texts),
        "temperature": 0,
        "top_p": 0.9,
        "max_tokens": 2048,
    }
    last_err = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            r = SESSION.post(LM_URL, json=payload, timeout=TIMEOUT_S)
            r.raise_for_status()
            raw = r.json()["choices"][0]["message"]["content"]
            return parse_entities(raw)
        except Exception as e:
            last_err = e
            if attempt < MAX_RETRIES:
                time.sleep(0.3)
    if len(texts) > 1:
        mid = len(texts) // 2
        r1 = call_llm(texts[:mid])
        r2 = call_llm(texts[mid:])
        return {k: r1.get(k, []) + r2.get(k, []) for k in GROUPS}
    print(f"    [skip] {texts[0][:60]!r} ({last_err})")
    return {k: [] for k in GROUPS}


# ==================== ZLICZANIE =========================

def count_occurrences(entity: str, all_texts: List[str]) -> int:
    """Ile komorek w Excelu zawiera ta encje (case-sensitive substring)."""
    return sum(1 for t in all_texts if entity in t)


# ==================== SAVE XLSX =========================

def save_entities_xlsx(
    entities: Dict[str, List[str]],
    pseudonyms: Dict[str, Dict[str, str]],   # {cat: {entity: pseudo}}
    counts: Dict[str, Dict[str, int]],        # {cat: {entity: count}}
    out_path: str,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Encje"

    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["Kategoria", "Wartosc", "Pseudonim", "Ile_razy", "Anonimizuj"]
    header_fill = PatternFill("solid", fgColor="2F3640")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[1].height = 22

    # Dropdown TAK/NIE
    dv = DataValidation(type="list", formula1='"TAK,NIE"', allow_blank=False, showDropDown=False)
    dv.sqref = "E2:E9999"
    ws.add_data_validation(dv)

    row = 2
    for cat, desc in GROUPS.items():
        vals = sorted(entities.get(cat, []), key=lambda v: counts.get(cat, {}).get(v, 0), reverse=True)
        if not vals:
            continue

        color = GROUP_COLORS.get(cat, "888888")
        cat_fill = PatternFill("solid", fgColor=color)
        cat_font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
        val_font = Font(name="Arial", size=9)
        pseudo_font = Font(name="Consolas", size=9, color="1565C0", bold=True)
        count_font = Font(name="Arial", size=9, color="666666")
        tak_fill = PatternFill("solid", fgColor="E8F5E9")
        tak_font = Font(name="Arial", size=9, color="2E7D32", bold=True)

        for val in vals:
            pseudo = pseudonyms.get(cat, {}).get(val, val)
            cnt = counts.get(cat, {}).get(val, 0)

            c_cat = ws.cell(row=row, column=1, value=cat)
            c_cat.font, c_cat.fill = cat_font, cat_fill
            c_cat.alignment = Alignment(horizontal="center", vertical="center")
            c_cat.border = border

            c_val = ws.cell(row=row, column=2, value=val)
            c_val.font = val_font
            c_val.alignment = Alignment(vertical="center")
            c_val.border = border

            c_pseudo = ws.cell(row=row, column=3, value=pseudo)
            c_pseudo.font = pseudo_font
            c_pseudo.alignment = Alignment(horizontal="center", vertical="center")
            c_pseudo.border = border

            c_cnt = ws.cell(row=row, column=4, value=cnt)
            c_cnt.font = count_font
            c_cnt.alignment = Alignment(horizontal="center", vertical="center")
            c_cnt.border = border

            c_anon = ws.cell(row=row, column=5, value="TAK")
            c_anon.font, c_anon.fill = tak_font, tak_fill
            c_anon.alignment = Alignment(horizontal="center", vertical="center")
            c_anon.border = border

            ws.row_dimensions[row].height = 18
            row += 1

    # Conditional formatting na kolumne E
    nie_fill = PatternFill("solid", fgColor="FFEBEE")
    nie_font = Font(name="Arial", size=9, color="C62828", bold=True)
    tak_fill2 = PatternFill("solid", fgColor="E8F5E9")
    tak_font2 = Font(name="Arial", size=9, color="2E7D32", bold=True)
    ws.conditional_formatting.add(
        f"E2:E{row}",
        CellIsRule(operator="equal", formula=['"NIE"'], fill=nie_fill, font=nie_font),
    )
    ws.conditional_formatting.add(
        f"E2:E{row}",
        CellIsRule(operator="equal", formula=['"TAK"'], fill=tak_fill2, font=tak_font2),
    )

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 14
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{row - 1}"
    wb.save(out_path)


# ==================== PIPELINE ==========================

def should_analyze(value) -> bool:
    if not isinstance(value, str):
        return False
    s = value.strip()
    if len(s) < MIN_LEN or s.startswith("="):
        return False
    if SKIP_RE.fullmatch(s):
        return False
    return True


def analyze_workbook(in_path: str) -> None:
    # ========== FAZA 1: Skan ==========
    print(f"[1/4] Skanuje (read-only): {in_path}")
    t0 = time.time()
    wb = load_workbook(in_path, read_only=True, data_only=True)
    print(f"       Zaladowano w {time.time()-t0:.1f}s. Arkusze: {wb.sheetnames}")

    all_texts: List[str] = []   # wszystkie oryginalne wartosci (do zliczania)
    cell_vals: List[str] = []   # oczyszczone (do LLM)
    total_cells = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        cnt = 0
        for row in ws.iter_rows():
            for c in row:
                total_cells += 1
                if should_analyze(c.value):
                    raw = str(c.value)
                    all_texts.append(raw)
                    cell_vals.append(clean_cell_value(raw))
                    cnt += 1
        if cnt:
            print(f"       {sn}: {cnt}")
    wb.close()

    unique = list(dict.fromkeys(cell_vals))
    print(f"       Razem: {len(cell_vals)} tekstowych, {len(unique)} unikalnych")

    # ========== FAZA 2: LLM ==========
    batches = [(i, unique[i:i+BATCH_SIZE]) for i in range(0, len(unique), BATCH_SIZE)]
    total_b = len(batches)
    print(f"[2/4] LLM: {total_b} batchy, {WORKERS} workery...")
    t1 = time.time()

    raw_entities: Dict[str, OrderedDict] = {k: OrderedDict() for k in GROUPS}
    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(call_llm, b): idx for idx, b in batches}
        for future in as_completed(futures):
            try:
                result = future.result()
                for cat, vals in result.items():
                    for v in vals:
                        if v and len(v) >= 2:
                            norm = normalize_entity(v)
                            if norm and len(norm) >= 2:
                                raw_entities[cat][norm] = True
            except Exception:
                pass
            done += 1
            if done % 5 == 0 or done == total_b:
                elapsed = time.time() - t1
                eta = int((elapsed / done) * (total_b - done))
                print(f"  {done}/{total_b} | {elapsed:.0f}s | ETA: {eta}s")

    entities = {k: list(v.keys()) for k, v in raw_entities.items()}
    total_found = sum(len(v) for v in entities.values())
    print(f"       Znaleziono {total_found} encji w {time.time()-t1:.1f}s")

    # ========== FAZA 3: Pseudonimy i zliczanie ==========
    print(f"[3/4] Generuje pseudonimy i zlicza wystapienia...")
    used_pseudonyms: set = set()
    pseudonyms: Dict[str, Dict[str, str]] = {}
    counts: Dict[str, Dict[str, int]] = {}

    for cat in GROUPS:
        pseudonyms[cat] = {}
        counts[cat] = {}
        for entity in entities.get(cat, []):
            pseudonyms[cat][entity] = make_pseudonym(entity, cat, used_pseudonyms)
            counts[cat][entity] = count_occurrences(entity, all_texts)

    for cat in GROUPS:
        n = len(entities.get(cat, []))
        if n:
            print(f"       {cat}: {n}")

    # ========== FAZA 4: Zapis ==========
    print(f"[4/4] Zapisuje: {ENTITIES_FILE}")
    save_entities_xlsx(entities, pseudonyms, counts, ENTITIES_FILE)

    print(f"\nGotowe! ({time.time()-t0:.1f}s)")
    print(f"  Plik do edycji: {ENTITIES_FILE}")
    print(f"  Kolumny:")
    print(f"    Wartosc    - co znaleziono")
    print(f"    Pseudonim  - propozycja zamiennika (mozesz zmienic)")
    print(f"    Ile_razy   - ile razy wystepuje w Excelu")
    print(f"    Anonimizuj - TAK/NIE")
    print(f"\n  Nastepnie: python anonymize.py {in_path}")


def find_inputs() -> List[str]:
    return [f for f in sorted(glob.glob("*.xlsx"))
            if not f.endswith("_anon.xlsx")
            and not f.startswith("~$")
            and f != ENTITIES_FILE]


def main() -> int:
    inputs = sys.argv[1:] if len(sys.argv) > 1 else find_inputs()
    if not inputs:
        print("Nie znaleziono .xlsx do analizy.")
        return 1
    try:
        SESSION.get(LM_URL.replace("/chat/completions", "/models"), timeout=5)
    except Exception as e:
        print(f"[UWAGA] LM Studio nie odpowiada: {e}")
        return 2
    analyze_workbook(inputs[0])
    return 0


if __name__ == "__main__":
    sys.exit(main())
