# analysis.py with multi-educator support
import pandas as pd
import json
from datetime import datetime
import unicodedata
import re


def normalize_name(text):
    """Normalize text: remove accents, extra spaces, convert to lowercase"""
    if not text:
        return ""
    # Remove accents
    text = "".join(
        c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
    )
    # Remove extra whitespace and lowercase
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def is_activity_cancelled(activity_text, desc_general, residents_text):
    """Check if activity is cancelled by looking for 'annul' variations in all fields"""
    # Normalize all texts
    activity_norm = normalize_name(activity_text)
    desc_norm = normalize_name(desc_general)
    residents_norm = normalize_name(residents_text)

    # Check for "annul" (catches annulé, annule, annulée, etc.)
    cancelled_pattern = r"annul"

    return (
        bool(re.search(cancelled_pattern, activity_norm))
        or bool(re.search(cancelled_pattern, desc_norm))
        or bool(re.search(cancelled_pattern, residents_norm))
    )


def load_employees():
    with open("employees.json", "r", encoding="utf-8") as f:
        return json.load(f)


def clean(text):
    if pd.isna(text):
        return ""
    return str(text).replace("\n", " ").strip()


def extract_all_educators_from_activity(text, employees):
    text_clean = clean(text)
    text_normalized = normalize_name(text_clean)

    educators = []
    for full_name in employees:
        name_normalized = normalize_name(full_name)
        if name_normalized in text_normalized:
            educators.append(full_name)
    return educators


def parse_resident_block(text):
    if pd.isna(text):
        lines = []
    else:
        lines = [l.strip() for l in str(text).split("\n") if l.strip()]

    residents = []
    for line in lines:
        if "a participé" in line:
            parts = line.split("a participé", 1)
            name = parts[0].strip()
            note = parts[1].replace(":", "").strip() if len(parts) > 1 else ""
            residents.append({"name": name, "status": "a participé", "note": note})
        else:
            residents.append({"name": line, "status": "", "note": ""})
    return residents


def analyze_excel(path, mode="hard"):
    df = pd.read_excel(path, header=None)
    employees = load_employees()

    activities = []
    today = datetime.now().date()

    for i in range(len(df)):
        date_raw = df.iloc[i, 0]
        try:
            date = pd.to_datetime(date_raw).date()
        except:
            continue

        if date >= today:
            break

        activity_block = clean(df.iloc[i, 1])
        desc_general = clean(df.iloc[i, 2])
        residents_block = df.iloc[i, 3]

        if not activity_block:
            continue
        if "appel" in activity_block.lower():
            continue

        educators = extract_all_educators_from_activity(activity_block, employees)

        residents = parse_resident_block(residents_block)

        # Check for cancelled using improved detection - check ALL columns in this row
        row_text = " ".join([clean(df.iloc[i, col]) for col in range(len(df.columns))])
        if is_activity_cancelled(row_text, "", ""):
            activities.append(
                {
                    "date_obj": date,
                    "date": date.strftime("%d/%m/%Y"),
                    "activity": activity_block,
                    "educators": educators,
                    "desc": desc_general,
                    "residents": residents,
                }
            )
            continue

        # Not cancelled - check for errors based on mode
        errors = []
        participated = any(r["status"].startswith("a participé") for r in residents)

        if mode == "soft":
            # Soft mode: only flag if no participation
            if not participated:
                activities.append(
                    {
                        "date_obj": date,
                        "date": date.strftime("%d/%m/%Y"),
                        "activity": activity_block,
                        "educators": educators,
                        "desc": desc_general,
                        "residents": residents,
                        "errors": ["Aucun résident n'a participé"],
                    }
                )
        else:
            # Hard mode: check both participation and descriptions
            if not participated:
                errors.append("Aucun résident n'a participé")
            elif desc_general and participated:
                for r in residents:
                    if r["status"].startswith("a participé") and not r["note"].strip():
                        errors.append(f"{r['name']} a participé sans note individuelle")
                        break

            if errors:
                activities.append(
                    {
                        "date_obj": date,
                        "date": date.strftime("%d/%m/%Y"),
                        "activity": activity_block,
                        "educators": educators,
                        "desc": desc_general,
                        "residents": residents,
                        "errors": errors,
                    }
                )

    # Sort by date (earliest first)
    activities.sort(key=lambda x: x["date_obj"], reverse=False)

    # Remove the temporary date_obj key
    for act in activities:
        del act["date_obj"]

    return activities, []
