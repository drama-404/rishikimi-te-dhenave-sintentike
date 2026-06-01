#!/usr/bin/env python3
"""
Synthetic survey data pipeline for:
AI adoption in Albanian banking services.

This script is designed for simulation, ML training, and EDA only. It writes
synthetic records with explicit provenance fields and can export a CSV that
matches the Microsoft Forms Excel export schema in this project.

Demographic targets are tuned for a realistic Albanian banking-customer sample,
not a uniform random sample of the whole population.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import random
import statistics
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable


LIKERT_TO_NUM = {
    "Nuk pajtohem aspak": 1,
    "Nuk pajtohem": 2,
    "Neutral": 3,
    "Pajtohem": 4,
    "Pajtohem plotesisht": 5,
}
NUM_TO_LIKERT = {value: key for key, value in LIKERT_TO_NUM.items()}

Q10_OPTIONS = [
    "Mobile banking (aplikacioni i bankes ne telefon)",
    "Internet banking (permes kompjuterit/browserit)",
    "Karta debiti/krediti per pagesa online",
    "E-wallet (EasyPay, M-Pay, E-Posta, etj.)",
    "Nuk perdor asnje sherbim digjital bankar",
]

EXPORT_HEADERS = [
    "Id",
    "Start time",
    "Completion time",
    "Email",
    "Name",
    "Duke klikuar me poshte, konfirmoj qe kam lexuar informacionin me siper dhe jap pelqimin tim per pjesemarrje ne kete studim.\n",
    "Gjinia juaj",
    "Mosha juaj",
    "Niveli me i larte i arsimit qe keni perfunduar",
    "Statusi juaj aktual i punesimit",
    "Ne cfare fushe punoni apo keni studiuar?",
    "Ku jetoni aktualisht?\n",
    "Cila eshte banka juaj kryesore? (ajo qe perdorni me shume)",
    "Sa kohe keni qe jeni klient/e e kesaj banke?",
    "Cilat sherbime bankare digjitale perdorni aktualisht?",
    "Sa shpesh i perdorni sherbimet bankare digjitale?",
    "Sa te kenaqur jeni me sherbimet digjitale te bankes suaj?",
    "A keni ndervepruar ndonjehere me nje chatbot (asistent virtual) ne faqen ose aplikacionin e bankes suaj?",
    "Sa pajtoheni me pohimet e meposhtme mbi inteligjencen artificiale (IA) ne banka?.Kam njohuri te mjaftueshme mbi inteligjencen artificiale dhe mundesite e saj",
    "Sa pajtoheni me pohimet e meposhtme mbi inteligjencen artificiale (IA) ne banka?.Besoj se IA mund te permiresoje cilesine e sherbimeve bankare",
    "Sa pajtoheni me pohimet e meposhtme mbi inteligjencen artificiale (IA) ne banka?.Do te ndihesha rehat nese nje sistem IA do te vleresonte kerkesen time per kredi",
    "Sa pajtoheni me pohimet e meposhtme mbi inteligjencen artificiale (IA) ne banka?.Do te preferoja te nderveproja me nje chatbot IA per pyetje te thjeshta bankare",
    "Sa pajtoheni me pohimet e meposhtme mbi inteligjencen artificiale (IA) ne banka?.Jam i/e shqetesuar per privatesine e te dhenave te mia nese banka perdor IA",
    "Sa pajtoheni me pohimet e meposhtme mbi inteligjencen artificiale (IA) ne banka?.Mendoj se IA do te zevendesoje shume nga punet aktuale ne banka",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Jam i gatshem te perdor nje aplikacion bankar me rekomandime te personalizuara nga IA",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Besimi im ne bankat shqiptare do te rritej nese IA do te zbulonte mashtrimet ne llogarite e mia me shpejt",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Do te pranoja keshillim financiar te gjeneruar nga IA nese do te ishte falas",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Transparenca mbi menyren se si IA merr vendime eshte e rendesishme per mua",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Do te preferoja nje banke qe perdor IA per sherbime me te shpejta ndaj nje banke qe nuk e perdor",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Do te ndihesha me i/e sigurt nese nje punonjes njerezor do te verifikonte vendimet e IA",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Ne pergjithesi, jam gati te provoj sherbime bankare qe perdorin inteligjencen artificiale",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Jam i gatshem te perdor nje aplikacion bankar me rekomandime te personalizuara nga IA1",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Besimi im ne bankat shqiptare do te rritej nese IA do te zbulonte mashtrimet ne llogarite e mia me shpejt1",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Do te pranoja keshillim financiar te gjeneruar nga IA nese do te ishte falas1",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Transparenca mbi menyren se si IA merr vendime eshte e rendesishme per mua1",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Do te preferoja nje banke qe perdor IA per sherbime me te shpejta ndaj nje banke qe nuk e perdor1",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Do te ndihesha me i/e sigurt nese nje punonjes njerezor do te verifikonte vendimet e IA1",
    "Sa pajtoheni me pohimet e meposhtme mbi gatishmerine tuaj per te perdorur IA ne banka?.Ne pergjithesi, jam gati te provoj sherbime bankare qe perdorin inteligjencen artificiale1",
    "Sa pajtoheni me pohimet e meposhtme mbi shqetesimet tuaja per perdorimin e IA ne banka?.Mungesa e njohurive te mia mbi IA me pengon te perdor sherbime te tilla",
    "Sa pajtoheni me pohimet e meposhtme mbi shqetesimet tuaja per perdorimin e IA ne banka?.Kam frike se te dhenat e mia personale mund te keqperdoren nese banka perdor IA",
    "Sa pajtoheni me pohimet e meposhtme mbi shqetesimet tuaja per perdorimin e IA ne banka?.Preferoj te flas me nje njeri, jo me nje robot, kur kam ceshtje bankare",
    "Sa pajtoheni me pohimet e meposhtme mbi shqetesimet tuaja per perdorimin e IA ne banka?.Nuk besoj se IA mund te marre vendime te sakta financiare",
    "Sa pajtoheni me pohimet e meposhtme mbi shqetesimet tuaja per perdorimin e IA ne banka?.Kostot e mundshme shtese per sherbimet me IA me shqetesojne",
    "A keni ndonje koment ose sugjerim shtese mbi perdorimin e inteligjences artificiale ne sherbimet bankare ne Shqiperi?",
]


BASE_WEIGHTS = {
    # Banking-customer demographic targets supplied for the synthetic sample.
    "gender": {
        "Femer": 0.48,
        "Mashkull": 0.50,
        "Preferoj te mos pergjigjem": 0.02,
    },
    "age": {
        "18-25": 0.25,
        "26-35": 0.30,
        "36-45": 0.22,
        "46-55": 0.15,
        "55+": 0.08,
    },
    "education": {
        "Arsim baze (9-vjecar)": 0.08,
        "Arsim i mesem (gjimnaz/profesional)": 0.32,
        "Bachelor": 0.38,
        "Master": 0.20,
        "Doktorature": 0.02,
    },
    "employment": {
        "I/e punesuar": 0.55,
        "I/e papune": 0.10,
        "Student/e": 0.12,
        "I/e vetepunesuar": 0.15,
        "I/e dale ne pension": 0.08,
    },
    "field": {
        "Teknologji / IT": 0.08,
        "Finance / Banke / Sigurime": 0.10,
        "Arsim": 0.12,
        "Shendetesi / Mjekesi": 0.09,
        "Ndertim / Inxhinieri": 0.11,
        "Tregti / Sherbime": 0.25,
        "Administrate publike": 0.15,
        "Bujqesi": 0.10,
    },
    "location": {
        "Tirane": 0.35,
        "Durres": 0.12,
        "Elbasan": 0.08,
        "Vlore": 0.08,
        "Shkoder": 0.08,
        "Korce": 0.060,
        "Qytet tjeter": 0.15,
        "Zone rurale": 0.08,
    },
    # Approximate market-share targets supplied for the synthetic sample.
    "bank": {
        "BKT (Banka Kombetare Tregtare)": 0.28,
        "Raiffeisen Bank": 0.24,
        "Credins Bank": 0.15,
        "Intesa Sanpaolo Bank": 0.10,
        "OTP Bank": 0.08,
        "Tirana Bank": 0.05,
        "Other": 0.10,
    },
    "tenure": {
        "Me pak se 1 vit": 0.12,
        "1-3 vjet": 0.28,
        "3-5 vjet": 0.25,
        "Mbi 5 vjet": 0.35,
    },
}


@dataclass(frozen=True)
class ArchetypeProfile:
    name: str
    target_share: float
    traits: dict[str, float]
    comment_templates: tuple[str, ...]


ARCHETYPES = {
    "Early Adopter": ArchetypeProfile(
        name="Early Adopter",
        target_share=0.30,
        traits={
            "digital_use": 4.7,
            "satisfaction": 4.3,
            "ai_knowledge": 4.1,
            "perceived_usefulness": 4.5,
            "readiness": 4.4,
            "privacy_concern": 3.3,
            "human_preference": 2.6,
            "trust_barrier": 2.8,
            "transparency_need": 4.2,
        },
        comment_templates=(
            "IA mund te rrise shpejtesine e sherbimeve bankare, sidomos per rekomandime dhe zbulim mashtrimesh.",
            "Do ta provoja nese banka e shpjegon qarte si ruhen te dhenat dhe si kontrollohen vendimet.",
        ),
    ),
    "Cautious User": ArchetypeProfile(
        name="Cautious User",
        target_share=0.40,
        traits={
            "digital_use": 3.5,
            "satisfaction": 3.5,
            "ai_knowledge": 3.2,
            "perceived_usefulness": 3.7,
            "readiness": 3.4,
            "privacy_concern": 4.7,
            "human_preference": 3.8,
            "trust_barrier": 4.2,
            "transparency_need": 4.8,
        },
        comment_templates=(
            "Jam i hapur ndaj IA, por privatesia dhe transparenca jane kushtet kryesore.",
            "Sherbimet me IA mund te ndihmojne, por vendimet e rendesishme duhet te verifikohen nga punonjesit e bankes.",
        ),
    ),
    "Traditionalist": ArchetypeProfile(
        name="Traditionalist",
        target_share=0.20,
        traits={
            "digital_use": 1.8,
            "satisfaction": 2.8,
            "ai_knowledge": 2.2,
            "perceived_usefulness": 2.6,
            "readiness": 2.3,
            "privacy_concern": 4.1,
            "human_preference": 4.7,
            "trust_barrier": 4.0,
            "transparency_need": 4.2,
        },
        comment_templates=(
            "Per ceshtje bankare preferoj te flas me nje person dhe jo me sistem automatik.",
            "IA mund te perdoret per pyetje te thjeshta, por jo per vendime financiare te rendesishme.",
        ),
    ),
    "Professional / Expert": ArchetypeProfile(
        name="Professional / Expert",
        target_share=0.10,
        traits={
            "digital_use": 4.2,
            "satisfaction": 3.8,
            "ai_knowledge": 4.6,
            "perceived_usefulness": 4.4,
            "readiness": 4.0,
            "privacy_concern": 4.0,
            "human_preference": 3.4,
            "trust_barrier": 3.6,
            "transparency_need": 4.9,
        },
        comment_templates=(
            "Potenciali eshte i madh, por bankat duhet te forcojne qeverisjen e te dhenave dhe kontrollin rregullator.",
            "IA duhet te zbatohet gradualisht, me auditim, transparence dhe pergjegjesi te qarte institucionale.",
        ),
    ),
}


def normalize_weights(weights: dict[str, float]) -> dict[str, float]:
    total = sum(max(value, 0.0) for value in weights.values())
    if total <= 0:
        raise ValueError("Weights must contain at least one positive value.")
    return {key: max(value, 0.0) / total for key, value in weights.items()}


def weighted_choice(rng: random.Random, weights: dict[str, float]) -> str:
    normalized = normalize_weights(weights)
    pick = rng.random()
    cumulative = 0.0
    last_key = next(iter(normalized))
    for key, weight in normalized.items():
        cumulative += weight
        last_key = key
        if pick <= cumulative:
            return key
    return last_key


def weighted_quota_values(n: int, weights: dict[str, float], rng: random.Random) -> list[str]:
    """Create a shuffled list whose counts match target weights as closely as possible."""
    normalized = normalize_weights(weights)
    raw_counts = {key: value * n for key, value in normalized.items()}
    counts = {key: int(math.floor(value)) for key, value in raw_counts.items()}
    remaining = n - sum(counts.values())
    remainders = sorted(
        ((raw_counts[key] - counts[key], key) for key in raw_counts),
        reverse=True,
    )
    for _, key in remainders[:remaining]:
        counts[key] += 1

    values: list[str] = []
    for key, count in counts.items():
        values.extend([key] * count)
    rng.shuffle(values)
    return values


def build_demographic_quotas(n: int, rng: random.Random) -> list[dict[str, str]]:
    # Age and education are intentionally NOT quota-locked here: they are sampled
    # per-archetype in demographic_answer so that realistic correlations emerge
    # (younger -> Early Adopter, older/Traditionalist -> lower education, etc.).
    # Their marginals stay close to target via the archetype mix and the
    # archetype-conditional weights, but are emergent rather than forced.
    quota_groups = {
        "Q2_gender": "gender",
        "Q7_location": "location",
        "Q8_bank": "bank",
    }
    pools = {
        answer_key: weighted_quota_values(n, BASE_WEIGHTS[weight_key], rng)
        for answer_key, weight_key in quota_groups.items()
    }
    return [
        {answer_key: values[index] for answer_key, values in pools.items()}
        for index in range(n)
    ]


def clamp_int(value: float, low: int = 1, high: int = 5) -> int:
    return max(low, min(high, int(round(value))))


def likert_value(rng: random.Random, target: float, spread: float = 0.72) -> int:
    return clamp_int(rng.gauss(target, spread))


def likert_label(value: int) -> str:
    return NUM_TO_LIKERT[clamp_int(value)]


def adjust_weights(base: dict[str, float], multipliers: dict[str, float]) -> dict[str, float]:
    adjusted = dict(base)
    for key, factor in multipliers.items():
        if key in adjusted:
            adjusted[key] *= factor
    return normalize_weights(adjusted)


# University-degree levels and the fields that, in practice, require one.
# Someone whose highest completed education is basic (9-year) or secondary
# (high school) cannot be working in these professions. Students are exempt,
# because Q6 asks about the field one works in *or has studied in*.
DEGREE_LEVELS = {"Bachelor", "Master", "Doktorature"}
DEGREE_REQUIRED_FIELDS = {
    "Finance / Banke / Sigurime",
    "Shendetesi / Mjekesi",
    "Arsim",
    "Teknologji / IT",
}

# Education by age band, reflecting real Albanian schooling history:
#  - Under 26: secondary school is compulsory, so NO basic-only (enforced hard).
#    Young cohorts are heavily degree-holding (~70% Bachelor+).
#  - 40s: a large share stopped at high school.
#  - 50+: the majority only completed basic (9-year) education.
AGE_EDUCATION = {
    "18-25": {
        "Arsim baze (9-vjecar)": 0.00,
        "Arsim i mesem (gjimnaz/profesional)": 0.30,
        "Bachelor": 0.50,
        "Master": 0.18,
        "Doktorature": 0.02,
    },
    "26-35": {
        "Arsim baze (9-vjecar)": 0.03,
        "Arsim i mesem (gjimnaz/profesional)": 0.24,
        "Bachelor": 0.48,
        "Master": 0.22,
        "Doktorature": 0.03,
    },
    "36-45": {
        "Arsim baze (9-vjecar)": 0.15,
        "Arsim i mesem (gjimnaz/profesional)": 0.40,
        "Bachelor": 0.30,
        "Master": 0.13,
        "Doktorature": 0.02,
    },
    "46-55": {
        "Arsim baze (9-vjecar)": 0.45,
        "Arsim i mesem (gjimnaz/profesional)": 0.35,
        "Bachelor": 0.15,
        "Master": 0.05,
        "Doktorature": 0.00,
    },
    "55+": {
        "Arsim baze (9-vjecar)": 0.70,
        "Arsim i mesem (gjimnaz/profesional)": 0.22,
        "Bachelor": 0.06,
        "Master": 0.02,
        "Doktorature": 0.00,
    },
}

# Bank tenure by age band: nobody opens an account before ~20, so the young
# cohort cannot have 5+ years of history.
AGE_TENURE = {
    "18-25": {
        "Me pak se 1 vit": 0.35,
        "1-3 vjet": 0.45,
        "3-5 vjet": 0.20,
        "Mbi 5 vjet": 0.00,
    },
    "26-35": {
        "Me pak se 1 vit": 0.12,
        "1-3 vjet": 0.30,
        "3-5 vjet": 0.30,
        "Mbi 5 vjet": 0.28,
    },
}
# 36+ falls back to BASE_WEIGHTS["tenure"] (long tenures dominate).


def demographic_answer(
    rng: random.Random,
    profile: ArchetypeProfile,
    group: str,
    education: str | None = None,
    employment: str | None = None,
    age: str | None = None,
) -> str:
    weights = BASE_WEIGHTS[group]

    if group == "age":
        if profile.name == "Early Adopter":
            weights = adjust_weights(weights, {"18-25": 2.0, "26-35": 1.7, "55+": 0.35})
        elif profile.name == "Traditionalist":
            weights = adjust_weights(weights, {"46-55": 1.4, "55+": 1.8, "18-25": 0.45, "26-35": 0.55})
        elif profile.name == "Professional / Expert":
            weights = adjust_weights(weights, {"26-35": 1.4, "36-45": 1.4, "55+": 0.55})

    if group == "education":
        # Age is the primary driver of education level. Archetype applies only a
        # mild nudge within what the age band realistically allows.
        weights = dict(AGE_EDUCATION.get(age, BASE_WEIGHTS["education"]))
        if profile.name == "Professional / Expert":
            weights = adjust_weights(weights, {"Master": 1.8, "Doktorature": 2.5, "Bachelor": 1.3, "Arsim i mesem (gjimnaz/profesional)": 0.5, "Arsim baze (9-vjecar)": 0.15})
        elif profile.name == "Early Adopter":
            weights = adjust_weights(weights, {"Bachelor": 1.3, "Master": 1.2, "Arsim baze (9-vjecar)": 0.5})
        elif profile.name == "Traditionalist":
            weights = adjust_weights(weights, {"Arsim baze (9-vjecar)": 1.6, "Arsim i mesem (gjimnaz/profesional)": 1.2, "Master": 0.4, "Doktorature": 0.2})
        # Hard legal rule: secondary school is compulsory, so under-26s can never
        # report basic-only education.
        if age == "18-25":
            weights = {k: v for k, v in weights.items() if k != "Arsim baze (9-vjecar)"}
        weights = normalize_weights(weights)

    if group == "tenure":
        # Young customers cannot have long bank histories (first account ~age 20).
        weights = AGE_TENURE.get(age, BASE_WEIGHTS["tenure"])

    if group == "employment":
        # Age-conditioned employment with HARD real-world constraints:
        #   - Students exist ONLY in the 18-25 band.
        #   - Retirees exist ONLY in the 55+ band (Albanian pension age is ~61 for
        #     women and ~65 for men, so nobody under ~55 can be retired).
        weights = dict(BASE_WEIGHTS["employment"])
        if age == "18-25":
            weights = adjust_weights(weights, {"Student/e": 3.0, "I/e vetepunesuar": 0.4})
        elif age == "55+":
            weights = adjust_weights(weights, {"I/e dale ne pension": 5.0, "I/e punesuar": 0.6})
        if profile.name == "Traditionalist":
            weights = adjust_weights(weights, {"I/e dale ne pension": 2.0, "Student/e": 0.4})
        elif profile.name == "Professional / Expert":
            weights = adjust_weights(weights, {"I/e punesuar": 1.5, "I/e vetepunesuar": 1.4, "I/e papune": 0.3})
        # Enforce the hard rules regardless of the nudges above.
        if age != "18-25":
            weights = {k: v for k, v in weights.items() if k != "Student/e"}
        if age != "55+":
            weights = {k: v for k, v in weights.items() if k != "I/e dale ne pension"}
        weights = normalize_weights(weights)

    if group == "field":
        if profile.name == "Professional / Expert":
            weights = adjust_weights(
                weights,
                {
                    "Teknologji / IT": 4.0,
                    "Finance / Banke / Sigurime": 4.5,
                    "Administrate publike": 1.4,
                    "Bujqesi": 0.15,
                },
            )
        elif profile.name == "Early Adopter":
            weights = adjust_weights(weights, {"Teknologji / IT": 2.0, "Finance / Banke / Sigurime": 1.6})

        # Consistency rule: no university degree -> cannot work in a profession
        # that requires one. Students are exempt (they may be studying it now).
        if education is not None and education not in DEGREE_LEVELS and employment != "Student/e":
            filtered = {k: v for k, v in weights.items() if k not in DEGREE_REQUIRED_FIELDS}
            if filtered:
                weights = normalize_weights(filtered)

    return weighted_choice(rng, weights)


def generate_q10_services(rng: random.Random, profile: ArchetypeProfile) -> list[str]:
    digital = profile.traits["digital_use"]
    if digital < 2.2 and rng.random() < 0.45:
        return ["Nuk perdor asnje sherbim digjital bankar"]

    probabilities = {
        "Mobile banking (aplikacioni i bankes ne telefon)": min(0.95, 0.15 + digital * 0.18),
        "Internet banking (permes kompjuterit/browserit)": min(0.80, 0.05 + digital * 0.14),
        "Karta debiti/krediti per pagesa online": min(0.88, 0.12 + digital * 0.15),
        "E-wallet (EasyPay, M-Pay, E-Posta, etj.)": min(0.55, 0.02 + digital * 0.10),
    }
    selected = [option for option, probability in probabilities.items() if rng.random() < probability]
    if not selected:
        selected = [rng.choice(Q10_OPTIONS[:3])]
    return selected


def digital_frequency(rng: random.Random, profile: ArchetypeProfile, services: list[str]) -> str:
    if "Nuk perdor asnje sherbim digjital bankar" in services:
        return "Rralle ose kurre"
    target = likert_value(rng, profile.traits["digital_use"], 0.65)
    return {
        5: "Cdo dite",
        4: "Disa here ne jave",
        3: "Nje here ne jave",
        2: "Disa here ne muaj",
        1: "Rralle ose kurre",
    }[target]


def chatbot_experience(rng: random.Random, profile: ArchetypeProfile) -> str:
    digital = profile.traits["digital_use"]
    probabilities = {
        "Po": max(0.05, min(0.72, digital * 0.13)),
        "Jo": max(0.20, 0.72 - digital * 0.08),
        "Nuk jam i/e sigurt": 0.16,
    }
    return weighted_choice(rng, probabilities)


def generate_likert_blocks(rng: random.Random, profile: ArchetypeProfile) -> dict[str, dict[str, int]]:
    traits = profile.traits

    # --- Latent respondent factors -------------------------------------------
    # Real survey constructs correlate because the same person carries a general
    # disposition across related items, not because the items are literally the
    # same number. We model two partly-opposed latent factors per respondent:
    #   pro_ai  -> general openness/positivity toward AI (loads on PU + readiness)
    #   concern -> general anxiety/skepticism (loads on barriers, mildly opposes pro_ai)
    # This produces a *moderate* PU<->readiness correlation and a *negative*
    # barrier<->readiness correlation (H1a positive, H1b negative) while keeping
    # plenty of independent item-level noise so the data looks naturally scattered.
    # A general pro-AI factor is shared across PU and readiness (gives a positive
    # but moderate H1a). Construct-specific biases (pu_bias / read_bias) add
    # independent variance so the two scales are clearly NOT the same variable --
    # this is what pulls the correlation down from ~0.85 into a realistic ~0.55.
    pro_ai = rng.gauss(0.0, 0.55)
    concern = rng.gauss(-0.45 * pro_ai, 0.65)
    pu_bias = rng.gauss(0.0, 0.60)
    read_bias = rng.gauss(0.0, 0.60)
    SPREAD = 0.90

    def item(target: float, *extra: float) -> int:
        return likert_value(rng, target + sum(extra), SPREAD)

    q14 = {
        "q14_1_ai_knowledge": item(traits["ai_knowledge"], 0.5 * pro_ai),
        "q14_2_service_quality": item(traits["perceived_usefulness"], pro_ai, pu_bias),
        "q14_3_ai_credit_eval": item(traits["perceived_usefulness"] - 0.35, pro_ai, pu_bias),
        "q14_4_chatbot_preference": item(traits["readiness"] - 0.10, pro_ai, read_bias),
        "q14_5_privacy_concern": item(traits["privacy_concern"], concern),
        "q14_6_job_replacement": item((traits["privacy_concern"] + traits["trust_barrier"]) / 2.0, 0.7 * concern),
    }
    q15 = {
        "q15_1_personalized_app": item(traits["readiness"], pro_ai, read_bias),
        "q15_2_fraud_detection_trust": item(traits["perceived_usefulness"] + 0.10, pro_ai, pu_bias),
        "q15_3_free_ai_advice": item(traits["readiness"] - 0.10, pro_ai, read_bias),
        "q15_4_transparency": item(traits["transparency_need"], 0.45 * concern),
        "q15_5_faster_ai_bank": item(traits["readiness"], pro_ai, read_bias),
        "q15_6_human_verification": item(traits["transparency_need"] - 0.10, 0.55 * concern),
        "q15_7_overall_readiness": item(traits["readiness"], pro_ai, read_bias),
    }
    # The real export has Q16 as a duplicate matrix. We keep it separate and
    # slightly jittered for schema compatibility, but omit it from core scores.
    q16 = {key.replace("q15", "q16"): likert_value(rng, value, 0.32) for key, value in q15.items()}
    q17 = {
        "q17_1_lack_knowledge": item(6.0 - traits["ai_knowledge"], 0.7 * concern),
        "q17_2_data_misuse": item(traits["privacy_concern"], concern),
        "q17_3_human_preference": item(traits["human_preference"], concern),
        "q17_4_inaccurate_decisions": item(traits["trust_barrier"], concern),
        "q17_5_extra_costs": item((traits["trust_barrier"] + traits["privacy_concern"]) / 2.0 - 0.15, 0.8 * concern),
    }
    return {"q14": q14, "q15": q15, "q16": q16, "q17": q17}


def calculate_scores(blocks: dict[str, dict[str, int]]) -> dict[str, float]:
    q14 = blocks["q14"]
    q15 = blocks["q15"]
    q17 = blocks["q17"]
    # PU and Readiness use DISJOINT item sets so H1a (PU -> Readiness) measures a
    # real relationship between two distinct constructs, not item overlap.
    # PU = belief that AI is beneficial; Readiness = willingness to actually use it.
    perceived_usefulness = mean_values(
        [
            q14["q14_2_service_quality"],
            q14["q14_3_ai_credit_eval"],
            q15["q15_2_fraud_detection_trust"],
        ]
    )
    adoption_readiness = mean_values(
        [
            q15["q15_1_personalized_app"],
            q15["q15_3_free_ai_advice"],
            q15["q15_5_faster_ai_bank"],
            q15["q15_7_overall_readiness"],
        ]
    )
    barriers = mean_values(
        [
            q14["q14_5_privacy_concern"],
            q17["q17_1_lack_knowledge"],
            q17["q17_2_data_misuse"],
            q17["q17_3_human_preference"],
            q17["q17_4_inaccurate_decisions"],
            q17["q17_5_extra_costs"],
        ]
    )
    trust_privacy_barrier = mean_values(
        [
            q14["q14_5_privacy_concern"],
            q17["q17_2_data_misuse"],
            q17["q17_4_inaccurate_decisions"],
        ]
    )
    governance_need = mean_values(
        [
            q15["q15_4_transparency"],
            q15["q15_6_human_verification"],
            q17["q17_2_data_misuse"],
            q17["q17_4_inaccurate_decisions"],
        ]
    )
    return {
        "perceived_usefulness": round(perceived_usefulness, 3),
        "adoption_readiness": round(adoption_readiness, 3),
        "barriers": round(barriers, 3),
        "trust_privacy_barrier": round(trust_privacy_barrier, 3),
        "governance_need": round(governance_need, 3),
    }


def mean_values(values: Iterable[float]) -> float:
    values = list(values)
    return sum(values) / len(values)


def archetype_distance(scores: dict[str, float], profile: ArchetypeProfile) -> float:
    expected = {
        "perceived_usefulness": profile.traits["perceived_usefulness"],
        "adoption_readiness": profile.traits["readiness"],
        "trust_privacy_barrier": mean_values([profile.traits["privacy_concern"], profile.traits["trust_barrier"]]),
        "governance_need": profile.traits["transparency_need"],
    }
    squared = [(scores[key] - expected[key]) ** 2 for key in expected]
    return math.sqrt(sum(squared) / len(squared))


def validate_record(record: dict[str, Any], profile: ArchetypeProfile) -> dict[str, Any]:
    scores = calculate_scores(record["answers"]["likert_blocks"])
    distance = archetype_distance(scores, profile)
    confidence = round(max(0.0, min(1.0, 1.0 - distance / 3.0)), 3)
    adjustments: list[str] = []
    blocks = record["answers"]["likert_blocks"]

    def nudge(block: str, key: str, by: int = 1) -> None:
        blocks[block][key] = clamp_int(blocks[block][key] + by)

    # Soft self-correction: only touch records that *clearly* contradict their
    # archetype, and nudge by one step rather than pinning to the extreme. This
    # keeps the archetype signal without creating artificial spikes at 4/5.
    if profile.name == "Cautious User" and scores["trust_privacy_barrier"] < 3.0:
        if blocks["q14"]["q14_5_privacy_concern"] <= 3:
            nudge("q14", "q14_5_privacy_concern")
        if blocks["q17"]["q17_2_data_misuse"] <= 3:
            nudge("q17", "q17_2_data_misuse")
        adjustments.append("Nudged privacy/trust concerns up for cautious profile.")
    elif profile.name == "Traditionalist" and blocks["q17"]["q17_3_human_preference"] <= 2:
        nudge("q17", "q17_3_human_preference")
        adjustments.append("Nudged human-contact preference up for traditionalist profile.")
    elif profile.name == "Early Adopter" and scores["adoption_readiness"] < 3.0:
        nudge("q15", "q15_7_overall_readiness")
        nudge("q15", "q15_5_faster_ai_bank")
        adjustments.append("Nudged readiness up for early adopter profile.")
    elif profile.name == "Professional / Expert" and scores["governance_need"] < 3.4:
        nudge("q15", "q15_4_transparency")
        nudge("q15", "q15_6_human_verification")
        adjustments.append("Nudged governance need up for expert profile.")

    scores = calculate_scores(record["answers"]["likert_blocks"])
    distance = archetype_distance(scores, profile)
    confidence = round(max(0.0, min(1.0, 1.0 - distance / 3.0)), 3)
    record["scores"] = scores
    record["validation"] = {
        "confidence": confidence,
        "archetype_distance": round(distance, 3),
        "adjustments": adjustments,
    }
    return record


class GeminiClient:
    def __init__(
        self,
        api_key: str | None,
        model: str = "gemini-2.5-flash-lite",
        rpm_delay_seconds: float = 4.2,
    ) -> None:
        self.api_key = api_key
        self.model = model
        self.rpm_delay_seconds = rpm_delay_seconds
        self.last_call_at = 0.0

    @property
    def enabled(self) -> bool:
        return bool(self.api_key)

    def generate_comment(self, record: dict[str, Any], profile: ArchetypeProfile) -> str | None:
        if not self.enabled:
            return None
        now = time.time()
        wait_for = self.rpm_delay_seconds - (now - self.last_call_at)
        if wait_for > 0:
            time.sleep(wait_for)

        prompt = {
            "task": "Generate one concise Albanian survey comment for synthetic test data.",
            "constraints": [
                "Return only JSON with key comment.",
                "The comment must be plausible for the archetype.",
                "Do not include personal data.",
                "Use plain Albanian without special formatting.",
                "Maximum 28 words.",
            ],
            "archetype": profile.name,
            "scores": record.get("scores", {}),
            "selected_answers": {
                "age": record["answers"]["Q3_age"],
                "field": record["answers"]["Q6_field"],
                "digital_frequency": record["answers"]["Q11_digital_frequency"],
            },
        }
        body = {
            "contents": [
                {
                    "role": "user",
                    "parts": [{"text": json.dumps(prompt, ensure_ascii=False)}],
                }
            ],
            "generationConfig": {
                "temperature": 0.75,
                "responseMimeType": "application/json",
            },
        }
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{self.model}:generateContent?key={self.api_key}"
        request = urllib.request.Request(
            url=url,
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                self.last_call_at = time.time()
                payload = json.loads(response.read().decode("utf-8"))
            text = payload["candidates"][0]["content"]["parts"][0]["text"]
            parsed = json.loads(text)
            comment = str(parsed.get("comment", "")).strip()
            return comment[:260] if comment else None
        except (KeyError, json.JSONDecodeError, urllib.error.URLError, TimeoutError):
            return None


def fallback_comment(rng: random.Random, profile: ArchetypeProfile) -> str:
    return rng.choice(profile.comment_templates)


# Fields associated with personal tech-savviness in the banking context.
TECH_SAVVY_FIELDS = {"Teknologji / IT", "Finance / Banke / Sigurime"}


def tech_savvy_plausible(age: str, education: str, field: str) -> bool:
    """Whether a respondent could realistically be a tech-savvy early adopter."""
    if age in ("18-25", "26-35"):
        return True  # digital natives: tech-savvy regardless of education
    if age in ("36-45", "46-55"):
        # Mid-age adopters only when well educated AND in a tech/finance field.
        return education in DEGREE_LEVELS and field in TECH_SAVVY_FIELDS
    return False  # 55+ : ~99% are not tech-savvy, so never an early adopter


def _skeptic_archetype(age: str, education: str, rng: random.Random) -> str:
    """Fallback non-tech-savvy archetype derived from the demographics."""
    older = age in ("46-55", "55+")
    low_education = education not in DEGREE_LEVELS
    if older and low_education:
        # Older + poorly educated -> mostly low-digital traditionalist.
        return "Traditionalist" if rng.random() < 0.7 else "Cautious User"
    if low_education and rng.random() < 0.3:
        return "Traditionalist"
    return "Cautious User"  # default: aware but skeptical, slow to try new things


def resolve_archetype(candidate: str, age: str, education: str, field: str, rng: random.Random) -> str:
    """Reassign tech-savvy archetypes when age/education/field make them implausible."""
    degree = education in DEGREE_LEVELS
    if candidate == "Early Adopter":
        if tech_savvy_plausible(age, education, field):
            return candidate
        return _skeptic_archetype(age, education, rng)
    if candidate == "Professional / Expert":
        # Experts are educated professionals in tech/finance, not the 55+ cohort.
        if degree and field in TECH_SAVVY_FIELDS and age != "55+":
            return candidate
        return _skeptic_archetype(age, education, rng)
    return candidate  # Cautious User / Traditionalist are always plausible


def generate_record(
    record_id: int,
    profile: ArchetypeProfile,
    rng: random.Random,
    llm: GeminiClient | None = None,
    demographic_quota: dict[str, str] | None = None,
) -> dict[str, Any]:
    start = datetime(2026, 5, 31, 8, 0, tzinfo=timezone.utc) + timedelta(minutes=record_id * 3 + rng.randint(0, 4))
    completion = start + timedelta(seconds=rng.randint(75, 420))
    quota = demographic_quota or {}

    # Resolve demographics first (the candidate profile provides only soft skews).
    # Age drives education and tenure; education+employment gate field selection.
    age = quota.get("Q3_age") or demographic_answer(rng, profile, "age")
    education = quota.get("Q4_education") or demographic_answer(rng, profile, "education", age=age)
    employment = demographic_answer(rng, profile, "employment", age=age)
    field = demographic_answer(rng, profile, "field", education=education, employment=employment)
    tenure = demographic_answer(rng, profile, "tenure", age=age)

    # Gate the archetype ON the demographics: a 55+ or poorly-educated respondent
    # cannot be a tech-savvy early adopter / expert. This may reassign the
    # archetype, after which all attitudinal answers follow the FINAL archetype.
    profile = ARCHETYPES[resolve_archetype(profile.name, age, education, field, rng)]
    q10 = generate_q10_services(rng, profile)
    likert_blocks = generate_likert_blocks(rng, profile)

    record = {
        "metadata": {
            "id": record_id,
            "synthetic": True,
            "purpose": "simulation/testing only for ML training and EDA",
            "archetype": profile.name,
            "generated_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        },
        "answers": {
            "start_time_iso": start.isoformat(timespec="seconds"),
            "completion_time_iso": completion.isoformat(timespec="seconds"),
            "Q1_consent": "Po, jap pelqimin tim",
            "Q2_gender": quota.get("Q2_gender", demographic_answer(rng, profile, "gender")),
            "Q3_age": age,
            "Q4_education": education,
            "Q5_employment": employment,
            "Q6_field": field,
            "Q7_location": quota.get("Q7_location", demographic_answer(rng, profile, "location")),
            "Q8_bank": quota.get("Q8_bank", demographic_answer(rng, profile, "bank")),
            "Q9_tenure": tenure,
            "Q10_digital_services": q10,
            "Q11_digital_frequency": digital_frequency(rng, profile, q10),
            "Q12_digital_satisfaction": likert_value(rng, profile.traits["satisfaction"]),
            "Q13_chatbot_interaction": chatbot_experience(rng, profile),
            "likert_blocks": likert_blocks,
            "Q18_comment": "",
        },
        "scores": {},
        "validation": {},
    }
    record = validate_record(record, profile)

    comment = llm.generate_comment(record, profile) if llm else None
    record["answers"]["Q18_comment"] = comment or fallback_comment(rng, profile)
    return record


def archetype_sequence(n: int, rng: random.Random) -> list[ArchetypeProfile]:
    raw_counts = {name: profile.target_share * n for name, profile in ARCHETYPES.items()}
    counts = {name: int(math.floor(value)) for name, value in raw_counts.items()}
    remaining = n - sum(counts.values())
    remainders = sorted(
        ((raw_counts[name] - counts[name], name) for name in raw_counts),
        reverse=True,
    )
    for _, name in remainders[:remaining]:
        counts[name] += 1
    sequence: list[ArchetypeProfile] = []
    for name, count in counts.items():
        sequence.extend([ARCHETYPES[name]] * count)
    rng.shuffle(sequence)
    return sequence


def generate_dataset(n: int, seed: int, llm: GeminiClient | None) -> list[dict[str, Any]]:
    rng = random.Random(seed)
    quotas = build_demographic_quotas(n, rng)
    return [
        generate_record(index + 1, profile, rng, llm, quotas[index])
        for index, profile in enumerate(archetype_sequence(n, rng))
    ]


def excel_serial(dt_iso: str) -> float:
    dt = datetime.fromisoformat(dt_iso)
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
    excel_epoch = datetime(1899, 12, 30)
    return (dt - excel_epoch).total_seconds() / 86400


def forms_export_row(record: dict[str, Any]) -> dict[str, Any]:
    answers = record["answers"]
    blocks = answers["likert_blocks"]
    q14 = blocks["q14"]
    q15 = blocks["q15"]
    q16 = blocks["q16"]
    q17 = blocks["q17"]
    values = [
        record["metadata"]["id"],
        excel_serial(answers["start_time_iso"]),
        excel_serial(answers["completion_time_iso"]),
        "anonymous",
        "",
        answers["Q1_consent"],
        answers["Q2_gender"],
        answers["Q3_age"],
        answers["Q4_education"],
        answers["Q5_employment"],
        answers["Q6_field"],
        answers["Q7_location"],
        answers["Q8_bank"],
        answers["Q9_tenure"],
        ";".join(answers["Q10_digital_services"]) + ";",
        answers["Q11_digital_frequency"],
        answers["Q12_digital_satisfaction"],
        answers["Q13_chatbot_interaction"],
        likert_label(q14["q14_1_ai_knowledge"]),
        likert_label(q14["q14_2_service_quality"]),
        likert_label(q14["q14_3_ai_credit_eval"]),
        likert_label(q14["q14_4_chatbot_preference"]),
        likert_label(q14["q14_5_privacy_concern"]),
        likert_label(q14["q14_6_job_replacement"]),
        likert_label(q15["q15_1_personalized_app"]),
        likert_label(q15["q15_2_fraud_detection_trust"]),
        likert_label(q15["q15_3_free_ai_advice"]),
        likert_label(q15["q15_4_transparency"]),
        likert_label(q15["q15_5_faster_ai_bank"]),
        likert_label(q15["q15_6_human_verification"]),
        likert_label(q15["q15_7_overall_readiness"]),
        likert_label(q16["q16_1_personalized_app"]),
        likert_label(q16["q16_2_fraud_detection_trust"]),
        likert_label(q16["q16_3_free_ai_advice"]),
        likert_label(q16["q16_4_transparency"]),
        likert_label(q16["q16_5_faster_ai_bank"]),
        likert_label(q16["q16_6_human_verification"]),
        likert_label(q16["q16_7_overall_readiness"]),
        likert_label(q17["q17_1_lack_knowledge"]),
        likert_label(q17["q17_2_data_misuse"]),
        likert_label(q17["q17_3_human_preference"]),
        likert_label(q17["q17_4_inaccurate_decisions"]),
        likert_label(q17["q17_5_extra_costs"]),
        answers["Q18_comment"],
    ]
    return dict(zip(EXPORT_HEADERS, values))


def analysis_row(record: dict[str, Any]) -> dict[str, Any]:
    answers = record["answers"]
    flat: dict[str, Any] = {
        "id": record["metadata"]["id"],
        "synthetic": record["metadata"]["synthetic"],
        "archetype": record["metadata"]["archetype"],
        "validation_confidence": record["validation"]["confidence"],
        "archetype_distance": record["validation"]["archetype_distance"],
        "gender": answers["Q2_gender"],
        "age": answers["Q3_age"],
        "education": answers["Q4_education"],
        "employment": answers["Q5_employment"],
        "field": answers["Q6_field"],
        "location": answers["Q7_location"],
        "bank": answers["Q8_bank"],
        "tenure": answers["Q9_tenure"],
        "digital_services": "|".join(answers["Q10_digital_services"]),
        "digital_frequency": answers["Q11_digital_frequency"],
        "digital_satisfaction": answers["Q12_digital_satisfaction"],
        "chatbot_interaction": answers["Q13_chatbot_interaction"],
        **record["scores"],
    }
    for block in answers["likert_blocks"].values():
        flat.update(block)
    flat["comment"] = answers["Q18_comment"]
    return flat


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        return
    fieldnames = fieldnames or list(rows[0].keys())
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def summarize(records: list[dict[str, Any]]) -> dict[str, Any]:
    def counts(key_path: tuple[str, ...]) -> dict[str, int]:
        result: dict[str, int] = {}
        for record in records:
            value: Any = record
            for key in key_path:
                value = value[key]
            result[str(value)] = result.get(str(value), 0) + 1
        return dict(sorted(result.items()))

    score_keys = records[0]["scores"].keys() if records else []
    return {
        "n": len(records),
        "synthetic_only": True,
        "archetype_counts": counts(("metadata", "archetype")),
        "gender_counts": counts(("answers", "Q2_gender")),
        "age_counts": counts(("answers", "Q3_age")),
        "education_counts": counts(("answers", "Q4_education")),
        "location_counts": counts(("answers", "Q7_location")),
        "bank_counts": counts(("answers", "Q8_bank")),
        "score_means": {
            key: round(statistics.mean(record["scores"][key] for record in records), 3)
            for key in score_keys
        },
        "validation_confidence_mean": round(
            statistics.mean(record["validation"]["confidence"] for record in records),
            3,
        )
        if records
        else None,
        "hypothesis_checks": hypothesis_checks(records),
    }


def pearson(xs: list[float], ys: list[float]) -> float | None:
    if len(xs) < 2 or len(xs) != len(ys):
        return None
    xbar = statistics.mean(xs)
    ybar = statistics.mean(ys)
    numerator = sum((x - xbar) * (y - ybar) for x, y in zip(xs, ys))
    xden = math.sqrt(sum((x - xbar) ** 2 for x in xs))
    yden = math.sqrt(sum((y - ybar) ** 2 for y in ys))
    if xden == 0 or yden == 0:
        return None
    return numerator / (xden * yden)


def hypothesis_checks(records: list[dict[str, Any]]) -> dict[str, Any]:
    if not records:
        return {}
    pu = [record["scores"]["perceived_usefulness"] for record in records]
    readiness = [record["scores"]["adoption_readiness"] for record in records]
    barriers = [record["scores"]["barriers"] for record in records]
    trust_privacy = [record["scores"]["trust_privacy_barrier"] for record in records]
    return {
        "H1a_pu_readiness_correlation": round(pearson(pu, readiness) or 0.0, 3),
        "H1b_trust_privacy_readiness_correlation": round(pearson(trust_privacy, readiness) or 0.0, 3),
        "H1b_trust_privacy_mean": round(statistics.mean(trust_privacy), 3),
        "overall_barrier_mean": round(statistics.mean(barriers), 3),
        "interpretation": (
            "Expected synthetic pattern: perceived usefulness correlates positively "
            "with readiness (H1a, moderate positive r), while trust/privacy barriers "
            "correlate negatively with readiness and stay among the strongest barriers (H1b)."
        ),
    }


def post_json(url: str, payload: Any, token: str | None = None) -> dict[str, Any]:
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    request = urllib.request.Request(
        url=url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        body = response.read().decode("utf-8")
        return {"status": response.status, "body": body[:500]}


def submit_to_form(records: list[dict[str, Any]], mode: str, batch_size: int = 25) -> dict[str, Any]:
    rows = [forms_export_row(record) for record in records]
    if mode == "dry_run":
        return {
            "mode": "dry_run",
            "submitted": 0,
            "message": "No network calls made. Export files are ready.",
        }

    if mode == "webhook":
        url = os.getenv("POWER_AUTOMATE_WEBHOOK_URL")
        if not url:
            raise RuntimeError("POWER_AUTOMATE_WEBHOOK_URL is required for webhook mode.")
        submitted = 0
        responses = []
        for batch in chunked(rows, batch_size):
            responses.append(post_json(url, {"rows": batch, "synthetic": True}))
            submitted += len(batch)
            time.sleep(0.5)
        return {"mode": mode, "submitted": submitted, "responses": responses}

    if mode == "graph_excel":
        token = os.getenv("MS_GRAPH_ACCESS_TOKEN")
        drive_item_id = os.getenv("EXCEL_WORKBOOK_DRIVE_ITEM_ID")
        table_name = os.getenv("EXCEL_TABLE_NAME", "Table1")
        if not token or not drive_item_id:
            raise RuntimeError("MS_GRAPH_ACCESS_TOKEN and EXCEL_WORKBOOK_DRIVE_ITEM_ID are required.")
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{drive_item_id}/workbook/tables/{table_name}/rows/add"
        submitted = 0
        responses = []
        for batch in chunked(rows, batch_size):
            values = [[row.get(header, "") for header in EXPORT_HEADERS] for row in batch]
            responses.append(post_json(url, {"values": values}, token=token))
            submitted += len(batch)
            time.sleep(1.0)
        return {
            "mode": mode,
            "submitted": submitted,
            "responses": responses,
            "warning": "Rows are appended to Excel, not submitted as native Microsoft Forms responses.",
        }

    raise ValueError(f"Unsupported submit mode: {mode}")


def chunked(items: list[Any], size: int) -> Iterable[list[Any]]:
    for index in range(0, len(items), size):
        yield items[index : index + size]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate synthetic Albanian banking AI survey data.")
    parser.add_argument("--n", type=int, default=400, help="Number of synthetic records to generate.")
    parser.add_argument("--seed", type=int, default=20260531, help="Random seed for reproducibility.")
    parser.add_argument("--outdir", default="output", help="Directory for generated CSV/JSON outputs.")
    parser.add_argument(
        "--use-llm",
        action="store_true",
        help="Use Gemini for Q18 comments when GEMINI_API_KEY is set. Falls back locally if unavailable.",
    )
    parser.add_argument(
        "--gemini-model",
        default=os.getenv("GEMINI_MODEL", "gemini-2.5-flash-lite"),
        help="Gemini model name. Default favors free-tier daily request capacity.",
    )
    parser.add_argument(
        "--submit-mode",
        choices=["dry_run", "webhook", "graph_excel"],
        default="dry_run",
        help="Submission route. dry_run never sends data.",
    )
    parser.add_argument("--batch-size", type=int, default=25, help="Submission batch size.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    outdir = Path(args.outdir)
    llm = None
    if args.use_llm:
        llm = GeminiClient(api_key=os.getenv("GEMINI_API_KEY"), model=args.gemini_model)
        if not llm.enabled:
            print("GEMINI_API_KEY is not set. Continuing with local comment templates.")

    records = generate_dataset(args.n, args.seed, llm)
    export_rows = [forms_export_row(record) for record in records]
    analysis_rows = [analysis_row(record) for record in records]
    summary = summarize(records)

    write_csv(outdir / "synthetic_forms_export.csv", export_rows, EXPORT_HEADERS)
    write_csv(outdir / "synthetic_analysis.csv", analysis_rows)
    write_json(outdir / "synthetic_records.json", records)
    write_json(outdir / "analysis_summary.json", summary)

    submission_result = submit_to_form(records, args.submit_mode, args.batch_size)
    write_json(outdir / "submission_result.json", submission_result)

    print(json.dumps({"outputs": str(outdir), "summary": summary, "submission": submission_result}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
