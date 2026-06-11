#!/usr/bin/env python3
"""
Result-calculation script for the compostable-packaging evidence corpus.

Design principle
----------------
Classification is organized by indicator, with each regex block tied to a specific coding purpose.
Each indicator has its own rule dictionary or named cue group:
source authority, stated organic-waste treatment route, acceptance, certification and approval basis,
rejection rationale, and the component fields used to build application-level compatibility scenarios.

Application-decision scenario logic is separated from reference-level
acceptance. Certification alone is treated as product qualification, while
official-only items, local/facility approval, positive-list inclusion, event or
dedicated routes, and other controlled-route restrictions remain scenario-relevant.

Inputs
------
- web_sources.zip containing 01_references_corpus.jsonl inside the archived source-corpus root folder

Outputs
-------
- results/display/: reader-facing CSV tables used for manuscript figures and Supplementary Data
- results/audit/: long-form regex/rule traceability CSVs and validation checks

Zip archives are no longer created by default. This script stops at the coding/output-table layer.
Excel workbook selection, sheet layout, path masking, column hiding, and formatting are handled by supplementary_data_generation.py.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import pandas as pd
    from pandas.api.types import is_object_dtype, is_string_dtype
except ImportError as exc:  # pragma: no cover
    raise SystemExit("This script requires pandas. Install with: pip install pandas") from exc

SCRIPT_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
INPUT_ZIP_DEFAULT = SCRIPT_DIR / "web_sources.zip"
OUTDIR_DEFAULT = SCRIPT_DIR / "results"
ZIP_OUTPUTS_DEFAULT = False

# ============================================================
# Core helpers
# ============================================================


def norm(s: Any) -> str:
    s = "" if s is None else str(s)
    s = s.replace("–", "-").replace("—", "-").replace("’", "'").replace("‘", "'")
    s = s.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def compile_rx(pattern: str) -> re.Pattern:
    return re.compile(pattern, re.I | re.S)


def first_match(patterns: Sequence[str], text: str) -> Tuple[str, str]:
    for pat in patterns:
        m = compile_rx(pat).search(text)
        if m:
            return pat, re.sub(r"\s+", " ", m.group(0)).strip()
    return "", ""


def has_any(patterns: Sequence[str], text: str) -> bool:
    return bool(first_match(patterns, text)[0])


def excerpt_around(text: str, match_text: str, width: int = 280) -> str:
    if not match_text:
        return ""
    nt = norm(text)
    nm = norm(match_text)
    pos = nt.find(nm[:80])
    if pos < 0:
        pos = nt.find(nm[:30])
    if pos < 0:
        return match_text[:width]
    s = max(0, pos - width // 2)
    e = min(len(text), pos + len(match_text) + width // 2)
    return re.sub(r"\s+", " ", text[s:e]).strip()


@dataclass(frozen=True)
class Rule:
    category: str
    include: Sequence[str]
    exclude: Sequence[str] = field(default_factory=list)
    definition: str = ""
    included_examples: str = ""
    excluded_examples: str = ""
    coding_basis: str = "regex_include_exclude"
    note: str = ""
    priority: int = 100


@dataclass
class RuleHit:
    category: str
    matched_rule: str = ""
    matched_text: str = ""
    excluded_by: str = ""
    coding_basis: str = ""
    excerpt: str = ""


NO_HIT = RuleHit(category="")


def apply_rules(text: str, rules: Sequence[Rule], default: str = "", *, source_text_for_excerpt: Optional[str] = None) -> RuleHit:
    """Apply include/exclude regex rules in priority order."""
    txt = norm(text)
    raw = text if source_text_for_excerpt is None else source_text_for_excerpt
    for rule in sorted(rules, key=lambda r: r.priority):
        ex_pat, ex_text = first_match(rule.exclude, txt) if rule.exclude else ("", "")
        if ex_pat:
            continue
        in_pat, in_text = first_match(rule.include, txt)
        if in_pat:
            return RuleHit(
                category=rule.category,
                matched_rule=in_pat,
                matched_text=in_text,
                coding_basis=rule.coding_basis,
                excerpt=excerpt_around(raw, in_text),
            )
    return RuleHit(category=default, coding_basis="fallback_default")


TARGET_TERMS = compile_rx(r"""
(compostable|compostability|composting|compost|biodegradable|biodegradability|bioplastic|bioplastics|bio[- ]?plastic|bio[- ]?based|
 organic\s*waste|organics|food\s*waste|biowaste|bio[- ]?waste|green\s*bin|brown\s*bin|food\s*scraps?|anaerobic\s*digestion|biogas|
 kompostier\w*|biologisch\s+abbaubar\w*|biokunststoff\w*|bioplastik|bioabfall|biom[üu]ll|biotonne|braune\s+tonne|
 composteer\w*|biologisch\s+afbreekbaar\w*|bioplastic\w*|gft|groente[-, ]*fruit[- ]*en[- ]*tuinafval|etensresten|bioafval|
 compostable\w*|biod[ée]gradable\w*|bioplastique\w*|d[ée]chets?\s+alimentaires?|biod[ée]chets?|bac\s+(brun|vert)|ordures?\s+organiques?|
 compostabil\w*|biodegradabil\w*|bioplastic\w*|rifiut[io]\s+organic[io]|umido|organico|contenitore\s+marrone|
 compostable\w*|biodegradable\w*|biopl[áa]stic\w*|residuos?\s+org[áa]nic\w*|biorresiduos?|contenedor\s+(marr[óo]n|verde)|org[aâ]nico|
 matavfall|bioavfall|bioaffald|madaffald|ruokaj[äa]te|bioj[äa]te|kompostoituva|biohajoava|bioj[äa]tepussi|
 bioodpad\w*|bioodpadu|kuchynsk\w*\s+bioodpad|kompostovatel\w*|biologicky\s+rozlo\w*|odpady\s+bio|pojemnik\s+bio|
 bioj[äa][äa]t\w*|biolagunev\w*|kompostitav\w*|pakend\w*|
 生ごみ|生ゴミ|堆肥|厨余|廚餘|有机垃圾|有機垃圾|餐厨|餐廚|易腐垃圾|垃圾分类|垃圾分類|可降解|可堆肥|塑料|塑膠|袋|プラスチック|음식물|음식물쓰레기|생분해|퇴비|비닐|플라스틱|봉투)
""",)


def evidence_windows(text: str, width: int = 700, max_windows: int = 90) -> str:
    """Return ranked evidence windows instead of the first keyword hits.

    The first-hit approach can over-sample navigation menus on long municipal
    pages. Here we still use multilingual target terms, but score each window
    by whether it contains operational acceptance/rejection language and
    material terms such as packaging, bags, liners, and compostable plastics.
    """
    raw = text or ""
    if not raw:
        return ""

    accept_reject_rx = compile_rx(r"accepted|allowed|permitted|can\s+(?:also\s+)?be\s+placed|can\s+go|may\s+go|put\s+in|place\s+in|only\s+use|must\s+be|should\s+be|not\s+accepted|not\s+allowed|do\s+not|cannot|prohibited|forbidden|rejected|contaminat|d[üu]rfen\s+nicht|darf\s+nicht|geh[öo]ren\s+nicht|nicht\s+in|kein\s+plastik|mag\s+niet|mogen\s+niet|niet\s+in|niet\s+bij|hoort\s+niet|toegestaan|mag\s+w[eé]l|pas\s+accept|interdit|refus|accept[ée]|autoris|no\s+se\s+aceptan|no\s+depositar|sin\s+(?:ning[uú]n\s+tipo\s+de\s+)?envase|se\s+aceptan|pueden\s+depositarse|vietat|si\s+possono\s+conferire|vanno\s+gettati|禁止|不可|不能|入れない|배출불가|금지")
    material_rx = compile_rx(r"compostable|biodegradable|bioplastic|bio[- ]?plastic|composteer|kompostier|biologisch\s+(?:afbreekbaar|abbaubar)|biod[ée]gradable|compostabil|biodegradabil|可堆肥|可降解|生分解|생분해|packag|emballage|envase|embalaje|verpakking|bag|liner|caddy|sac|zak|beutel|t[üu]te|bolsa|sacco|sacchetti|pussi|pose|worki|袋|봉투")
    organics_rx = compile_rx(r"brown\s+bin|green\s+bin|organics?|organic\s+waste|food\s+waste|food\s+scraps?|biowaste|bio[- ]?waste|gft|biotonne|bac\s+brun|contenedor\s+marr[oó]n|umido|bioj[äa]te|bioodpad|生ごみ|음식물")

    candidates: list[tuple[int, int, str]] = []
    for idx, m in enumerate(TARGET_TERMS.finditer(raw)):
        if idx > 1500:  # protect against pathological navigation-heavy pages
            break
        s = max(0, m.start() - width)
        e = min(len(raw), m.end() + width)
        snippet = re.sub(r"\s+", " ", raw[s:e]).strip()
        if not snippet:
            continue
        score = 10
        if material_rx.search(snippet):
            score -= 3
        if accept_reject_rx.search(snippet):
            score -= 4
        if organics_rx.search(snippet):
            score -= 2
        # De-prioritise pure navigation/menu fragments.
        low = norm(snippet[:500])
        if low.count("overview") + low.count("übersicht") + low.count("menu") + low.count("navigation") > 3:
            score += 3
        candidates.append((score, m.start(), snippet))

    if not candidates:
        return raw[:14000]

    candidates.sort(key=lambda x: (x[0], x[1]))
    selected: list[str] = []
    seen: set[str] = set()
    for _, _, snip in candidates:
        key = norm(snip)[:180]
        if key in seen:
            continue
        selected.append(snip)
        seen.add(key)
        if len(selected) >= max_windows:
            break

    head = re.sub(r"\s+", " ", raw[:1800]).strip()
    if head:
        selected.append(head)
    return "\n---WINDOW---\n".join(selected)


# ============================================================
# Transparent rule dictionaries
# ============================================================

SOURCE_AUTHORITY_RULES: List[Rule] = [
    Rule(
        "Supranational framework",
        [r"\b(european union|european commission|eu policy|eur-lex|oecd|united nations|unep|international)\b"],
        definition="Supranational law, policy, or guidance relevant across multiple countries.",
        included_examples="European Commission, EU framework, EUR-Lex, OECD, UN.",
        excluded_examples="National ministries implementing EU law; municipal or operator pages citing EU rules.",
        coding_basis="metadata_regex + fallback_from_source_label",
        priority=10,
    ),
    Rule(
        "National rule or policy",
        [
            r"\b(national|federal|ministry|ministerio|ministere|minist[eè]re|government|agency|department|epa|environment agency|parliament|law|act|regulation|ordinance|decree|bioabfv|ley\s+7/2022|boe|legifrance|gazette|wrap|recycle\s*now|vang-hha|rijkswaterstaat|mohu)\b",
            r"\b(national\s+government|central\s+government|official\s+national|country[- ]?wide)\b",
        ],
        definition="National-level law, policy, ministry guidance, or government-backed national sorting guidance.",
        included_examples="national government, federal ordinance, ministry, national environment agency, national law, national public guidance.",
        excluded_examples="State/provincial rules, city council pages, facility operator guidance.",
        coding_basis="metadata_regex + fallback_from_source_label",
        priority=20,
    ),
    Rule(
        "State, provincial, regional, or canton rule",
        [
            r"\b(state|province|provincial|regional|region|canton|county|prefecture|autonomous|landkreis|d[ée]partement|departement|flanders|wallonia|catalonia|california|ontario|quebec|québec|nova\s+scotia|alberta|busan\s+metropolitan|kumamoto\s+prefecture|kanagawa\s+prefecture)\b",
            r"\b(region|regional|county|prefecture)\b.{0,80}\b(council|authority|government|guidance|rule|waste)\b",
        ],
        definition="Subnational rule or guidance issued at state, province, region, canton, county, district, or prefecture level.",
        included_examples="state agency, province, canton, county/district authority, autonomous region, prefecture.",
        excluded_examples="Intermunicipal waste authorities or specific municipal sorting pages.",
        coding_basis="metadata_regex + fallback_from_source_label",
        priority=30,
    ),
    Rule(
        "Intermunicipal or regional waste authority",
        [
            r"\b(intermunicipal|metropolitan|regional waste|solid waste authority|waste authority|joint waste|m[eé]tropole|r[ée]gie|lipor|hsy|rrfb|metropole|metropolitana|syndicat|tratolixo|hvc\s+service|rad\s*/|rd4|saver\s+service|greater\s+manchester)\b",
        ],
        definition="Authority or partnership operating across several municipalities or a metropolitan service area.",
        included_examples="metropolitan waste authority, regional waste partnership, intermunicipal authority, multi-municipality utility.",
        excluded_examples="Single city council; single private waste company/facility operator.",
        coding_basis="metadata_regex + fallback_from_source_label",
        priority=40,
    ),
    Rule(
        "Waste operator or treatment-facility rule",
        [
            r"\b(facility|plant|operator|hauler|collector|processor|composter|composting facility|anaerobic|biogas|waste[- ]company|waste company|public utility|recology|wm\b|awm\b|awb\b|bsr\b|stadtreinigung|milieu|hvc|rd4|gesenu|contarina|salerno pulita|mohu|tratolixo|mustankorkea|biorepack|suez|veolia|envar|mywaste|pražské\s+služby|mpo\b|olo\s+a\.s\.|jp\s+voka\s+snaga|vasa\b|getliņi|getlini|voka\s+snaga|intradel|hygea|komunala\s+kranj|minett\s+kompost|calgary|halifax|reykjav[ií]k)\b",
        ],
        definition="Waste operator, collector, processor, composting facility, AD/biogas plant, public utility, or EPR/operator guidance.",
        included_examples="Recology, Stadtreinigung, composting facility, biogas operator, waste-company guidance, municipal utility operator.",
        excluded_examples="Municipal council webpages unless the source is explicitly the operator or utility.",
        coding_basis="metadata_regex + fallback_from_source_label",
        priority=50,
    ),
    Rule(
        "Municipal sorting rule",
        [
            r"\b(city|municipal|municipality|commune|gemeente|gemeinde|borough|district|council|town|ville|stadt|ayuntamiento|comune|cidade|municipio|município|municipalidad|ciudad|miasto|urząd|alcald[ií]a|oslo|berlin|munich|hamburg|cologne|amsterdam|vienna|milan|barcelona|vigo|toronto|seattle|austin|san francisco|tokyo|seoul|prague|alimos|riga|krak[oó]w|toru[nń]|warsaw|bratislava|vilnius|ljubljana|maribor|christchurch|auckland|vitacura|las\s+palmas|brno|copenhagen|reykjav[ií]k|portland|new\s+york\s+city|minamata|zushi|providencia|santa\s+juana|wrocław|wroclaw)\b",
        ],
        definition="Municipal/city-level sorting rule or local government guidance.",
        included_examples="city council, municipality, borough, district, gemeente, comune, ayuntamiento, municipal utility page.",
        excluded_examples="Regional/state/national frameworks; facility-only rules.",
        coding_basis="metadata_regex + fallback_from_source_label",
        priority=60,
    ),
]

TREATMENT_ROUTE_RULES: List[Rule] = [
    Rule("AD + composting", [r"\b(anaerobic\s+digestion|\bAD\b|biogas|biomethane|methani[sz]ation|m[ée]thanisation|biom[ée]thanisation|verg[äa]rung|vergisting|r[öo]tning|digestate|biog[öo]dsel|biog[øo]dning)\b.{0,180}\b(compost|composting|kompost|compostaggio|compostaje|compostagem|kompostering|kompostointi|kompostiranje|κομποστοποίηση|堆肥|퇴비)", r"\b(compost|composting|kompost|compostaggio|compostaje|compostagem|kompostering|kompostointi|kompostiranje|κομποστοποίηση|堆肥|퇴비)\b.{0,180}\b(anaerobic\s+digestion|\bAD\b|biogas|biomethane|methani[sz]ation|m[ée]thanisation|biom[ée]thanisation|verg[äa]rung|vergisting|r[öo]tning|digestate|biog[öo]dsel|biog[øo]dning)\b"], definition="The source states or strongly implies AD/biogas combined with or followed by composting.", priority=10),
    Rule("AD / biogas", [r"\b(anaerobic\s+digestion|\bAD\b|biogas|biomethane|methani[sz]ation|m[ée]thanisation|biom[ée]thanisation|verg[äa]rung|vergisting|biokaas|biogaz|biogass|biogasanlegg|r[öo]tning|biog[öo]dsel|biog[øo]dning|digestate|厌氧|厭氧|沼气|沼氣)\b"], definition="The stated organics route is anaerobic digestion, biogas, biomethane production, or digestate management.", priority=20),
    Rule("Composting", [r"\b(composting|compost\s+facility|industrial\s+compost|in[- ]?vessel\s+compost|IVC|compostage|compostaggio|compostaje|compostagem|kompostierung|kompostering|kompostointi|kompostiranje|composteerinstallatie|composteerbaar|kompostownia|κομποστοποίηση|κομπόστ|堆肥化|堆肥|퇴비화|퇴비|kompostitav\w*)\b"], definition="The stated organics route is composting without a clear AD element.", priority=30),
    Rule("Other valorisation", [r"\b(animal\s+feed|feedstock|valorisation|valorization|rendering|black\s+soldier\s+fly|insect|food\s+waste\s+recycling|resource\s+recovery|자원화|사료화)\b"], definition="The source indicates an organics recovery route other than composting or AD, such as animal feed, insect treatment, rendering, or other food-waste recycling.", priority=40),
    Rule("Treatment route not stated", [r"\b(brown\s*bin|green\s*cart|organic\s*waste\s*collection|food\s*waste\s*collection|food\s*scraps?\s*(?:bin|collection)|bio[- ]?waste\s*collection|biowaste\s*collection|FOGO|GFT|biotonne|bioabfall|contenedor\s+marr[oó]n|contentor\s+castanho|bac\s+brun|rifiut[io]\s+organic[io]|umido|maisto\s+atliek|bioj[äa]te|madaffald|matavfall|bioodpad|pojemnik\s+bio|kitchen\s*waste\s*collection)\b"], definition="The source does not state a downstream treatment technology such as composting, AD/biogas, or other organic recovery; this includes sources that state only the collection/bin route.", priority=45),
]

CERTIFICATION_RULES: List[Rule] = [
    Rule("OK compost HOME / NF T 51-800 / AS 5810", [r"\bok\s*compost\s*home\b", r"\bhome[- ]?compostable|home\s*compostability\b", r"\bnf\s*t\s*51[- ]?800\b", r"\bas\s*5810\b", r"\bcompost\s*home\b"], definition="Home-compostability standards and labels are mentioned.", priority=10),
    Rule("BPI / ASTM / CMA", [r"\bbpi\b|\bbiodegradable\s+products\s+institute\b", r"\bastm\s*d\s*6400\b|\bd\s*6400\b", r"\bastm\s*d\s*6868\b|\bd\s*6868\b", r"\b(?:cma[- ]?(?:certified|accepted|approved)|compost\s+manufacturing\s+alliance)\b"], definition="North American industrial compostability certification or acceptance programmes are mentioned.", priority=20),
    Rule("EN 13432 / OK compost / Seedling", [r"\b(en|din\s*en|uni\s*en|bs\s*en|nf\s*en)?\s*13432\b", r"\beuropean\s+standard\s+en\s*13432\b", r"\bok\s*compost(?!\s*home)\b", r"\bseedling\b|\bkeimling\b|\bkiemplant(?:logo)?\b", r"\bt[üu]v\s*(austria)?\b", r"\bdin\s*certco\b", r"\b(compostabile\s+cic|cic\s+compostabile)\b"], definition="European/industrial compostability standards and labels are mentioned.", priority=30),
    Rule("Country-specific standards/certifications (BNQ / AS 4736 / DINplus)", [r"\bas\s*4736\b", r"\bbnq\b|\bcan/bnq\b", r"\bdinplus\b", r"\bnmx\b|\bks\s*[a-z]?\s*\d{4}|\bjis\b", r"\bnational\s+(standard|certification)\b.{0,120}\b(compost|biodegrad)", r"\b(as\s*4736|as\s*5810)\s*:\s*\d{4}\b"], definition="Country-specific compostability standards, certification schemes, or stricter national add-ons are mentioned.", priority=40),
    Rule("Government/programme-approved", [r"\b(official|approved|designated|provided|supplied|distributed|council[- ]?supplied|municipal[- ]?supplied|city[- ]?supplied|programme[- ]?approved|program[- ]?approved)\b.{0,120}\b(bag|liner|sack|item|product)s?\b", r"\b(grid[- ]?print|lattice[- ]?print|printed)\s*bio\s*bag\b", r"(専用袋|指定袋|市が回収|전용봉투|종량제봉투|专用袋|專用袋|orange\s+bag|oranžin\w*\s+maišel\w*)"], definition="Acceptance is tied to an official, designated, positive-list, or programme-approved item rather than generic certification alone.", priority=50),
    Rule("Generic compostable/biodegradable claim only", [r"\b(compostable|biodegradable|industrially\s+compostable|certified\s+compostable|composteerbaar\w*|kompostierbar\w*|compostabil\w*|biod[ée]gradable\w*|biohajoava|kompostoituva|biolagunev\w*|biolo[šs]ko\s+razgradljiv\w*|kompostowaln\w*|biodegradowaln\w*)\b", r"(可堆肥|可生物降解|可降解|生分解|생분해)"], definition="Only a generic compostable/biodegradable claim is present; no named standard or approval basis is detected.", priority=90),
]

# Reusable text-pattern groups for acceptance, conditions, rationales, and applications.
ACCEPT_CUES = [
    r"\b(accepted|allowed|permitted|can\s+go|may\s+go|suitable|put\s+in|goes\s+in|include[d]?)\b.{0,180}\b(compostable|biodegradable|bioplastic|packaging|serviceware|cup|plate|container|cutlery|bag|liner)\b.{0,180}\b(organic|food\s*waste|green\s*bin|brown\s*bin|compost|biowaste|organics)",
    r"\b(compostable|biodegradable|bioplastic|packaging|serviceware|cup|plate|container|cutlery|bag|liner)\w*\b.{0,180}\b(accepted|allowed|permitted|can\s+go|may\s+go|suitable|put\s+in|goes\s+in|include[d]?)\b.{0,180}\b(organic|food\s*waste|green\s*bin|brown\s*bin|compost|biowaste|organics)",
    r"compostable\s+packaging\s*\(all\s+must\s+be\s+certified\s+to\s+en\s*13432\)",
    r"compostable\s+packaging\s+and\s+take\s*away\s+items.{0,180}can\s+also\s+be\s+placed\s+in\s+the\s+brown\s+bin",
    r"(composteerbare?\w*|biologisch\s+afbreekbare?\w*|biozak\w*|zakjes?|kiemplant|ok\s*compost).{0,180}(toegestaan|mag(?:\s+wel)?|wel\s+in|gooit|gooien|gebruiken|in\s+(de\s+)?groene\s+container|in\s+(de\s+)?gft)",
    r"(vanno\s+gettati|va\s+gettato|conferire|si\s+possono\s+conferire|sono\s+ammessi|pu[oò]\s+conferire).{0,180}(contenitore\s+marrone|umido|organico).{0,180}(compostabil\w*|bioplastic\w*|imballagg\w*|stoviglie|posate|sacchetti|shopper)",
    r"(accept[ée]s?|autoris[ée]s?|peuvent\s+(?:être\s+)?(?:mis|jet[ée]s?)).{0,180}(sacs?|emballages?|compostable\w*|biod[ée]gradable\w*).{0,180}(bac|biod[ée]chets?|d[ée]chets?\s+alimentaires|organiques?)",
    r"(se\s+aceptan|permitid[oa]s?|pueden\s+(?:ir|depositarse)|se\s+pueden\s+depositar).{0,180}(bolsas?|envases?|compostable\w*|biodegradable\w*).{0,180}(contenedor\s+marr[óo]n|biorresiduos?|residuos?\s+org[áa]nicos?)",
    r"(komposterbar\w*|bioposer|biohajoava\w*|kompostoituva\w*|compost\s*home).{0,180}(matavfall|bioavfall|bioj[äa]te|ruokaj[äa]te|ker[äa]ysastia|biojäteastiaan)",
]
LINER_ACCEPT_CUES = [
    r"\b(use|can\s+use|may\s+use|allowed\s+to\s+use|line|pack|collect|wrap)\b.{0,160}\b(compostable|bioplastic|biodegradable|paper)\s+(bag|liner|sack)s?",
    r"\b(food[- ]?waste|bio[- ]?waste|organic\s*waste|kitchen\s*caddy)\b.{0,160}\b(compostable|biodegradable|paper)\s+(bag|liner|sack)s?",
    r"\b(compostable|biodegradable|paper)\s+(bag|liner|sack)s?\b.{0,160}\b(food[- ]?waste|bio[- ]?waste|organic\s*waste|kitchen\s*caddy|brown\s+bin)",
    r"(bioj[äa]te|matavfall|food\s*waste|gft|gfe\+?t).{0,180}(compostable|composteerbare?\w*|biologisch\s+afbreekbare?\w*|biohajoava\w*|komposterbar\w*|paper|papieren?)\s+(bag|liner|pussi|zak|zakje|pose)s?",
    r"si\s+possono\s+utilizzare\s+solo\s+sacchetti\s+biodegradabili\s+compostabili",
    r"hormis\s+les\s+sacs\s+compostables\s+servant\s+[àa]\s+jeter\s+vos\s+d[ée]chets\s+alimentaires",
]
BROAD_ACCEPT_CUES = [
    # Broad acceptance must be an operational disposal instruction, not a
    # general statement that compostable packaging exists in the organic fraction.
    r"\bcompostable\s+packaging\b.{0,180}\b(accepted|allowed|permitted|can\s+(?:also\s+)?be\s+placed|can\s+go|may\s+go|put\s+in|place\s+in)\b.{0,120}\b(brown\s+bin|green\s+bin|organic\s+waste|organics|biowaste|food\s+waste|compost(?:ing)?\s+bin)",
    r"\b(accepted|allowed|permitted|can\s+(?:also\s+)?be\s+placed|can\s+go|may\s+go|put\s+in|place\s+in)\b.{0,180}\b(compostable\s+packaging|compostable\s+take[- ]?away|compostable\s+serviceware|compostable\s+food[- ]?service|compostable\s+(cups?|plates?|bowls?|containers?|cutlery|trays?))",
    r"compostable\s+packaging\s*\(all\s+must\s+be\s+certified\s+to\s+en\s*13432\)",
    r"compostable\s+packaging\s+and\s+take\s*away\s+items.{0,180}can\s+also\s+be\s+placed\s+in\s+the\s+brown\s+bin",
    r"\b(certified\s+compostable|bpi[- ]?certified)\b.{0,120}\b(cups?|plates?|bowls?|containers?|cutlery|serviceware|packaging)\b.{0,160}\b(accepted|allowed|permitted|green\s+bin|brown\s+bin|organics\s+bin)",
]
REJECT_CUES = [
    r"\b(not\s+accepted|not\s+allowed|do\s+not\s+put|don't\s+put|must\s+not|should\s+not|prohibited|rejected|not\s+permitted|not\s+suitable|cannot|can\s*not|can’t|can't|contamination|contaminant)\b.{0,180}\b(compostable|biodegradable|bioplastic|bio[- ]?plastic|compostable\s+packag|biodegradable\s+packag|serviceware|cup|cutlery)",
    r"\b(compostable|biodegradable|bioplastic|bio[- ]?plastic|compostable\s+packag|biodegradable\s+packag|serviceware|cup|cutlery)\w*\b.{0,200}\b(not\s+accepted|not\s+allowed|do\s+not\s+put|don't\s+put|must\s+not|should\s+not|prohibited|rejected|not\s+permitted|not\s+suitable|cannot|can\s*not|can’t|can't|contamination|contaminant)\b",
    r"(plastik|kunststoff|bioplastik|biokunststoff|kompostierbare?\w*\s+(biobeutel|beutel|t[üu]ten)|verpackungen|folien|t[üu]ten).{0,180}(d[üu]rfen\s+nicht|geh[öo]ren\s+nicht|nicht\s+in\s+(die\s+)?biotonne|darf\s+nicht\s+rein|kein\s+plastik|verboten|nicht\s+ausreichend)",
    r"(niet\s+in\s+(de\s+)?gft|mag\s+niet\s+(in|bij)|niet\s+bij\s+(het\s+)?gft|gooi\s+nooit|hoort\s+niet\s+in|niet\s+toegestaan).{0,200}(bioplastic|composteerbaar\w*|biologisch\s+afbreekbaar\w*|verpakking\w*|plastic|wegwerpbestek|borden|bekers|zakjes?)",
    r"(ne\s+(?:les?\s+)?mettez\s+pas|ne\s+pas\s+mettre|pas\s+dans|interdit\w*|refus[ée]s?|non\s+accept[ée]s?).{0,200}(compostable\w*|biod[ée]gradable\w*|bioplastique\w*|emballage\w*|plastique\w*)",
    r"(non\s+(?:vanno|va|mettere|ammessi|accettati|conferire)|vietat\w+|non\s+conferire).{0,200}(compostabil\w*|bioplastic\w*|imballagg\w*|plastica|sacchetti|stoviglie|posate)",
    r"(no\s+(?:se\s+)?(?:aceptan|permiten)|no\s+(?:va|van)|no\s+depositar|prohibid\w+).{0,200}(compostable\w*|biodegradable\w*|biopl[áa]stic\w*|envase\w*|pl[áa]stico\w*|bolsas?)",
    r"(ikke|inte|må\s+ikke|skal\s+ikke|ei|ei\s+saa|ei\s+kuulu).{0,180}(komposterbar\w*|bioposer|bioplast|plast|biohajoava\w*|kompostoituva\w*|muovipakkauk\w*|kertak[äa]ytt[öo]astia)",
    r"(可堆肥|可降解|生分解|塑料|塑膠|プラスチック|생분해|플라스틱|비닐).{0,100}(不(?:可|能|得)|不能|不可|請勿|不要|禁止|一般ごみ|燃えるごみ|종량제|일반쓰레기|배출불가)",
    r"emballages\s+m[eê]me\s+compostables\s+et/ou\s+biod[ée]gradables\s+ne\s+sont\s+pas\s+accept[ée]s",
    r"any\s+plastic\s+bags\s+and\s+liners,\s+including\s+compostable\s+bags",
]
POSITIVE_LIST_CUES = [r"\b(positive\s+list|approved\s+(product|item|list)|accepted\s+product\s+list|whitelist|named\s+items?|specific(?:ally)?\s+approved|only\s+(listed|named|specified|approved)|not\s+on\s+the\s+list|non\s+list[ée]s?)\b", r"\b(only|solo|seulement|nur|alleen)\b.{0,120}\b(approved|listed|named|specific|official|designated|autoris[ée]s?|zugelassen|goedgekeurd)"]
CONTROLLED_CUES = [r"\b(closed[- ]?loop|event|festival|venue|stadium|commercial\s+collection|controlled\s+(system|collection)|dedicated\s+collection|back[- ]?of[- ]?house|food[- ]?service|monocharge)\b"]
NO_ROUTE_CUES = [r"\b(no|not|without|lack(?:ing)?)\b.{0,120}\b(organic\s*waste|food\s*waste|biowaste|bio[- ]?waste|organics)\b.{0,120}\b(collection|route|service|infrastructure|facility|treatment|system)\b", r"\b(no|not|without|lack(?:ing)?)\b.{0,120}\b(industrial\s+composting|composting\s+facility|anaerobic\s+digestion|ad\s+facility)\b", r"large[- ]?scale\s+composting\s+facilities.{0,120}do\s+not\s+yet\s+exist"]
FRAGMENTED_CUES = [r"\b(var(?:y|ies)|various|depends\s+on|local\s+(rules|conditions|authority)|municipalities\s+decide|case[- ]?by[- ]?case|fragmented|unclear|not\s+clear|not\s+stated|varieert|verschilt|localement|commune\s+par\s+commune)\b"]

# Indicator: Acceptance and Application Decision - food-waste liner cues.
# These multilingual patterns capture explicit permission to use food-waste
# liners/collection bags, or explicit exceptions for liners in sources that
# otherwise reject compostable packaging.
LINER_ACCEPTANCE_CONTEXT_CUES = [
    # English: explicit liner permission or exception.
    r"\b(food[- ]?waste|kitchen|caddy|organics?|organic[- ]?waste|bio[- ]?waste)\s+(caddy\s+)?(bag|liner|sack)s?\b.{0,140}\b(can|may|must|should|accepted|allowed|permitted|collected|used|use)\b",
    r"\b(can|may|must|should|accepted|allowed|permitted|collected|used|use)\b.{0,120}\b(compostable|biodegradable|paper|kraft)\s+(food[- ]?waste\s+)?(bag|liner|sack)s?\b.{0,120}\b(food[- ]?waste|organic\s*waste|organics?|biowaste|green\s*cart|green\s*bin|brown\s*bin|caddy)\b",
    r"\b(other\s+than|except|exception|excluding|apart\s+from)\b.{0,80}\b(kitchen\s+caddy|food[- ]?waste|organics?|organic[- ]?waste)\s+(bag|liner|sack)s?",
    r"\b(if\s+bags\s+are\s+used|bags\s+used\s+for\s+food\s+waste).{0,120}\b(must|should|need\s+to)\b.{0,80}\b(paper|biodegradable|compostable)\s+(bag|liner|sack)s?",

    # French: food-waste/biowaste sacks accepted or explicit exception for such sacks.
    r"\b(hormis|sauf|except[ée]?|exception).{0,100}sacs?\s+(?:compostables?|biod[ée]gradables?|en\s+papier).{0,120}(d[ée]chets?\s+alimentaires|biod[ée]chets?|mati[èe]res?\s+organiques?)",
    r"sacs?\s+(?:compostables?|biod[ée]gradables?|bio|en\s+papier|kraft).{0,120}\b(?:accept[ée]s?|autoris[ée]s?|admis|utilis[ée]s?|peuvent\s+[êe]tre\s+utilis[ée]s?|servant\s+[àa]\s+jeter)\b.{0,120}(d[ée]chets?\s+alimentaires|biod[ée]chets?|bac\s+brun|d[ée]chets?\s+organiques)",
    r"(?:d[ée]chets?\s+alimentaires|biod[ée]chets?|d[ée]chets?\s+organiques|mati[èe]res?\s+organiques).{0,140}sacs?\s+(?:compostables?|biod[ée]gradables?|bio|en\s+papier|kraft)",

    # Spanish / Portuguese: compostable bags as the container/liner for organic waste.
    r"bolsas?\s+(?:compostables?|biodegradables?).{0,120}\b(?:se\s+(?:pueden|deben)\s+usar|se\s+aceptan|depositar|recoger|usar|utilizar|recomendable|obligatorio)\b.{0,140}(org[áa]nic[ao]s?|biorresiduos?|residuos?\s+alimentarios|contenedor\s+marr[oó]n)",
    r"sacos?\s+compost[áa]veis.{0,120}\b(?:podem\s+ser|devem\s+ser|aceites?|utilizados?)\b.{0,140}(res[íi]duos?\s+alimentares|org[âa]nicos|contentor\s+castanho)",

    # Italian: compostable sacks/shoppers explicitly placed in the wet organic fraction.
    r"(?:sacchetti|shopper|sacchi)\s+(?:biodegradabili\s+e\s+)?compostabili.{0,140}\b(?:vanno\s+gettati|si\s+possono\s+conferire|conferire|nell['’]?umido|contenitore\s+marrone|organico)\b",

    # Dutch: explicit positive wording or an exception for special GFT/GFE(T) bags.
    r"(?:uitzondering|speciaal\s+gft[- ]?zakje|gft[- ]?zakje|gfe\+?t[- ]?zakje|biozakje|composteerbaar\w*\s+zakje).{0,160}\b(?:mag\s+w[eé]l|mag|toegestaan|gebruiken|gebruik|in\s+(?:de\s+)?gft|in\s+(?:de\s+)?groene\s+container)\b",
    r"(?:kiemplantlogo|ok\s*compostlogo).{0,120}\b(?:zakje|zak|liner).{0,120}\b(?:mag|toegestaan|gebruiken|in\s+(?:de\s+)?gft)\b",

    # Nordic/Finnish/Baltic: explicit biowaste/food-waste bag collection wording.
    r"(?:matavfall|madaffald|bioaffald|bioavfall|bioj[äa]te|ruokaj[äa]te).{0,100}(?:bioposer|matavfallsp[åa]sar|bioavfallsp[åa]sar|bioj[äa]tepussi|bioj[äa]tepussit|paperipussi).{0,120}(?:skal|kan|må|saa|accepted|sorteres|anv[aä]ndas|bruges|käyttää)",
    r"(?:bioposer|matavfallsp[åa]sar|bioavfallsp[åa]sar|bioj[äa]tepussi|bioj[äa]tepussit).{0,120}(?:matavfall|madaffald|bioaffald|bioavfall|bioj[äa]te|ruokaj[äa]te)",

    # Slavic languages: compostable/biodegradable bag terms plus explicit use for kitchen/biowaste.
    r"(?:kompostovatel\w*|biolo[šs]ko\s+razgradljiv\w*|biologicky\s+rozlo[žz]iteln\w*|biodegradowaln\w*|kompostowaln\w*).{0,80}(?:vre[čc]k\w*|vrec[úu]ska|s[áa][čc]ky|worki).{0,140}(?:bioodpad|kuchynsk\w*\s+bioodpad|hned[áa]\s+n[áa]doba|rjavi\s+zabojnik|pojemnik\s+bio|bio)",
    r"(?:bioodpady?|pojemnik\s+bio|br[ąa]zowy\s+pojemnik|kuchynsk\w*\s+bioodpad).{0,140}(?:worki|bioworki|vre[čc]k\w*|vrec[úu]ska|s[áa][čc]ky).{0,80}(?:biodegradowaln\w*|kompostowaln\w*|kompostovatel\w*|rozlo[žz]iteln\w*)",

    # East Asian sources: require explicit biodegradable/compostable bag wording, not merely a designated municipal bag.
    r"(?:生ごみ|食品廃棄物).{0,140}(?:生分\w{0,8}解性\w{0,12}(?:プラスチック)?\w{0,8}袋|堆肥化可能\w*袋)",
    r"(?:음식물쓰레기|음식물류|음식물\s*폐기물).{0,100}(?:생분해\w*\s*봉투|퇴비화\w*\s*봉투)",
    r"(?:厨余垃圾|廚餘垃圾|餐厨垃圾|有机垃圾|有機垃圾).{0,100}(?:可堆肥袋|可降解袋|可生物降解袋)",
    r"(?:zul[äa]ssigen?|zertifizierten?)\s+(?:biobeutel|beutel).{0,160}(?:din\s*en\s*13432|keimling|din[- ]?plus|bioabfall)",
    r"biobeutel\s+aus\s+papier\s+und\s+biologisch\s+abbaubare\s+kunststoffbeutel.{0,180}(?:anforderungen|erfüllen|bioabfall)",

    # Additional observed corpus phrases after context review.
    r"gft[- ]?zakken?.{0,120}(?:toegelaten|aanbieden|gebruiken|in\s+de\s+container|mag(?:en)?)",
    r"(?:se\s+entregan|se\s+repartir[áa]n|repartir[áa]|entregar[áa]n).{0,120}bolsas?\s+compostables?.{0,160}(?:residuos?\s+org[áa]nicos|materia\s+org[áa]nica|contenedor\s+marr[oó]n)",
    r"bolsas?\s+compostables?.{0,160}(?:permitidas?|autorizadas?|listado|planta\s+de\s+disposici[oó]n|residuos?\s+org[áa]nicos|materia\s+org[áa]nica)",
    r"sacs?\s+(?:100%\s+)?biod[ée]gradables?.{0,160}(?:logo|minett\s+kompost|poubelle\s+bio|d[ée]chets?\s+organiques|compostage)",

]

BROAD_ACCEPTANCE_CONTEXT_CUES = [
    # Direct operational acceptance of compostable packaging/serviceware beyond liners.
    r"\b(compostable|certified\s+compostable|bpi[- ]?certified)\b.{0,100}\b(serviceware|food[- ]?service|take[- ]?away|cups?|plates?|bowls?|containers?|cutlery|packaging)\b.{0,140}\b(accepted|allowed|permitted|can\s+(?:go|be\s+placed)|may\s+(?:go|be\s+placed)|place[d]?\s+in|put\s+in)\b.{0,100}\b(green\s+bin|brown\s+bin|organics?|organic\s+waste|food\s+waste|compost(?:ing)?)\b",
    r"\b(compostable\s+packaging\s+and\s+take[- ]?away\s+items).{0,160}\b(can\s+also\s+be\s+placed\s+in\s+the\s+brown\s+bin)\b",
    r"(?:imballagg\w*|stoviglie|posate|bicchieri|piatti).{0,100}compostabil\w*.{0,140}(?:sono\s+ammessi|si\s+possono\s+conferire|vanno\s+gettati|contenitore\s+marrone|nell['’]?umido|organico)",
    r"(?:envases?|embalagens?|emballages?).{0,100}(?:compostables?|compost[áa]veis|biod[ée]gradables?).{0,140}(?:se\s+aceptan|permitid\w*|pueden\s+(?:depositarse|ir)|podem\s+ser\s+(?:colocados|aceites?)|accept[ée]s?).{0,100}(?:bac|contenedor|contentor|biorresiduos?|org[áa]nic)",
]

REJECTION_CONTEXT_CUES = [
    r"(emballages?.{0,80}(m[eê]me\s+)?(compostables?|biod[ée]gradables?).{0,120}(ne\s+sont\s+pas\s+accept[ée]s|interdits?|refus[ée]s?)|aucun\s+emballage|sans\s+emballage)",
    r"(sin\s+ning[uú]n\s+tipo\s+de\s+envase|sin\s+envase|sem\s+embalagem|retirar\s+.*envase|retirar\s+.*embalagem)",
    r"(niet\s+in\s+(de\s+)?gft|hoort\s+niet\s+in|gooi\s+nooit).{0,220}(composteerbaar\w*|biologisch\s+afbreekbaar\w*|verpakking\w*|plastic|zakjes?)",
    r"(kompostierbare?\w*\s+(biobeutel|beutel|t[üu]ten)|bioplastik|biokunststoff|verpackungen).{0,180}(d[üu]rfen\s+nicht|nicht\s+in\s+die\s+biotonne|kein\s+plastik|geh[öo]ren\s+nicht)",
    r"(opakowania|embala[žz]a|obaly|foli[oó]wki|plastikow\w*|plastov\w*).{0,220}(nie|ne|niso|nepat[řr][íi]|prepoved|zakaz)",
    r"(förpackningar|emballasje|emballage|pakkaukset|pakend|maisto\s+pakuot\w*).{0,180}(inte|ikke|må\s+ikke|ei|ei\s+saa|ei\s+kuulu|not\s+allowed)",
    r"(レジ袋|プラスチック袋|プラスチック|容器包装).{0,100}(二重袋|使わない|入れない|禁止|燃えるごみ|可燃ごみ)",
    r"(비닐|플라스틱|생분해|퇴비화).{0,120}(일반쓰레기|배출불가|금지|넣지|불가)",
    r"(可堆肥|可生物降解|可降解|塑料袋|塑膠袋|包装|包裝).{0,100}(不可|不能|請勿|请勿|不要|禁止|其他垃圾)",
    r"(πλαστικ[έε]ς?\s+σακούλες?|πλαστικ[άα]|συσκευασ[ίι]ες?).{0,140}(δεν|μην|απαγορε[ύυ]ονται|όχι)",
]

# Additional rejection phrases observed in facility/operator sorting lists.
REJECTION_CONTEXT_CUES += [
    r"(?:intrus\s+du\s+bac\s+brun|d[ée]chets?\s+interdits?).{0,120}(?:plastique\s+compostable|sacs?\s+ou\s+vaisselle/emballages|vaisselle/emballages)",
]



# Round-2 context review additions: patterns derived from corpus-wide review of
# remaining "No explicit rule" rows. These are not citation-specific overrides;
# they encode recurring operational wording observed across Dutch, Italian,
# Spanish/Portuguese, Greek, Lithuanian, Polish, Hungarian, and US sources.
LINER_ACCEPTANCE_MULTILINGUAL_CUES = [
    # Italian: brown-bin/organic-waste collection bags only.
    r"si\s+possono\s+utilizzare\s+solo\s+sacchetti\s+biodegradabili\s+compostabili",
    r"(?:contenitore\s+marrone|rifiuto\s+organico|raccolta\s+(?:dei\s+)?rifiuti\s+umidi).{0,160}sacchetti\s+biodegradabili\s+compostabili",

    # Spanish: organics must be deposited in compostable bags.
    r"residuos\s+(?:org[áa]nicos|de\s+procedencia\s+dom[ée]stica).{0,180}deben\s+depositarse\s+en\s+bolsas\s+(?:de\s+basura\s+)?compostables",
    r"(?:restos|residuos)\s+org[áa]nicos.{0,160}siempre\s+introducidos\s+en\s+bolsas\s+compostables",
    r"bolsas\s+compostables.{0,100}(?:fracci[óo]n\s+org[áa]nica|contenedor(?:es)?\s+marr[oó]n(?:es)?)",

    # Greek: food leftovers can be placed loose or in biodegradable/compostable/paper bags.
    r"(?:υπολείμματα\s+τροφ|τροφών).{0,180}(?:σακούλες).{0,120}(?:βιοαποδομήσιμες|κομποστοποιήσιμες|χάρτινες)",
    r"(?:βιοαποδομήσιμες|κομποστοποιήσιμες).{0,80}σακούλες.{0,120}(?:καφέ\s+κάδ|υπολείμματα\s+τροφ|τροφών)",

    # Spanish/Chilean liner exception: organics may be put in plastic or compostable bags,
    # while other biodegradable plastic items remain excluded.
    r"residuos\s+pueden\s+ir\s+en\s+bolsa\s+(?:pl[áa]stica\s+o\s+)?compostable",

    # Polish liner exception: only bags marked for composting are allowed.
    r"wyj[ąa]tkiem\s+s[ąa]\s+tylko\s+te,?\s+kt[óo]re\s+oznaczono\s+symbolem\s+do\s+kompostowania",

    # Catalan/Girona: compostable bags for selective collection of organic waste.
    r"bosses\s+compostables\s+per\s+a\s+la\s+recollida\s+selectiva\s+de\s+residus\s+org[àa]nics",
]

BROAD_ACCEPTANCE_MULTILINGUAL_CUES = [
    # Italian municipal positive lists: compostable coffee capsules, serviceware,
    # shoppers and soiled cardboard/paper packaging in the wet-organic stream.
    r"(?:umido|organico|scarti\s+dei\s+cibi).{0,240}(?:capsule\s+del\s+caff[èe]\s+biodegradabili\s+e\s+compostabili|piatti,\s*bicchieri\s+e\s+posate\s+biodegradabili\s+e\s+compostabili|shopper\s+biodegradabili\s+e\s+compostabili)",
    r"(?:piatti,\s*bicchieri\s+e\s+posate\s+biodegradabili\s+e\s+compostabili|capsule\s+del\s+caff[èe]\s+biodegradabili\s+e\s+compostabili).{0,220}(?:marchio\s+ok\s+compost|compostabile\s+cic|umido|organico)",

    # US/Recology-style green-cart lists with BPI compostable bags and natural-fiber serviceware.
    r"acceptable\s+composting\s+materials.{0,700}(?:b\.?p\.?i\.?\s+compostable\s+bags|bpi[- ]?certified).{0,500}(?:natural\s+fiber[- ]based\s+cups|plates|bowls|utensils|bagasse)",
    r"(?:b\.?p\.?i\.?\s+compostable\s+bags|bpi[- ]?certified).{0,300}(?:natural\s+fiber[- ]based\s+cups|plates|bowls|utensils|bagasse).{0,300}(?:green\s+composting\s+cart|compost(?:ing)?\s+materials)",

    # Mexico City standard: approved compostable single-use plastic bags/products
    # carry the disposal instruction "deposítese con la fracción orgánica".
    r"(?:bolsas\s+y\s+productos\s+pl[áa]sticos\s+de\s+un\s+solo\s+uso|productos\s+pl[áa]sticos).{0,220}compostables?.{0,220}depos[íi]tese\s+con\s+la\s+fracci[óo]n\s+org[áa]nica",
    r"depos[íi]tese\s+con\s+la\s+fracci[óo]n\s+org[áa]nica.{0,220}(?:logo\s+compostable|productos\s+compostables?|bolsas\s+compostables?)",
    r"(?:bags\s+and\s+single[- ]use\s+plastic\s+products|compostable\s+single[- ]use\s+plastic\s+bags\s+and\s+products).{0,420}depos[íi]tese\s+con\s+la\s+fracci[óo]n\s+org[áa]nica",
    r"depos[íi]tese\s+con\s+la\s+fracci[óo]n\s+org[áa]nica.{0,220}compostable\s+single[- ]use\s+plastic\s+bags\s+and\s+products",

    # Lithuanian/Alytus municipal list: biodegradable coffee capsules and wooden
    # disposable cutlery appear in the food/kitchen-waste accepted list.
    r"(?:maisto\s+ir\s+virtuv[ėe]s\s+atliekos|maisto\s+atliek).{0,700}(?:biologiškai\s+skaidžių\s+medžiagų\s+pagamintos\s+kavos\s+kapsulės|mediniai,\s*vienkartiniai\s+stalo\s+įrankiai)",
]

REJECTION_MULTILINGUAL_CUES = [
    # Dutch VANG yes/no GFT list: packaging and disposable products of any material;
    # compostable coffee cups are specifically excluded.
    r"niet\s+verpakkingen\s+en\s+wegwerpproducten\s+van\s+wat\s+voor\s+materiaal\s+dan\s+ook",
    r"koffiecups,?\s+ook\s+composteerbare",

    # UK food-waste caddy pages that reject compostable bags because they clog
    # machinery and are not composted.
    r"please\s+do\s+not\s+use:.{0,260}compostable\s+bags.{0,220}(?:clog\s+up\s+the\s+machinery|aren[’']?t\s+composted|sent\s+to\s+be\s+incinerated)",
    r"compostable\s+bags.{0,160}(?:break\s+down\s+too\s+quickly|clog\s+up\s+the\s+machinery|aren[’']?t\s+composted)",

    # Portuguese and Spanish organic-fraction rules excluding packaging/plastics/cups/cutlery.
    r"(?:biorres[íi]duos|sacos?\s+verdes?|contentor(?:es)?\s+(?:castanho|marr[oó]n)|contenedor(?:es)?\s+marr[oó]n).{0,260}(?:embalagens?\s+alimentares|envases?\s+de\s+(?:papel|cart[óo]n|vidrio|pl[áa]stico)|pl[áa]sticos?|metais|copos?,\s*talheres\s+e\s+loi[çc]as)",
    r"(?:no\s+debemos\s+tirar|no\s+deben\s+depositarse|tampoco\s+deben\s+depositarse).{0,260}(?:envases?|pl[áa]stico|latas|briks|aceite|pa[ñn]ales)",
    r"(?:neste\s+contentor\s+n[ãa]o\s+deve\s+colocar|por\s+outro\s+lado,?\s+neste\s+contentor\s+n[ãa]o\s+deve\s+colocar).{0,260}(?:embalagens?|pl[áa]stico|metal|vidro|cart[ãa]o|papel)",

    # Costa Rica / composting-facility guidance: plastic bags are opened and removed.
    r"(?:bolsas\s+pl[áa]sticas).{0,120}(?:deben\s+abrirse|vaciarse|retirar\s+el\s+pl[áa]stico)",

    # Chile Providencia: biodegradable plastics explicitly listed as not recyclable/accepted.
    r"pl[áa]sticos\s+biodegradables.{0,160}(?:bolsas\s+ca[ñn]a\s+de\s+az[úu]car|cubiertos|vasos|no\s+se\s+puede\s+reciclar)",

    # Polish Wroclaw: biodegradable bags do not work, except those marked compostable.
    # This encodes a restriction; the separate exception can still yield liner-only if
    # the positive marked-compostable-bag context is strong.
    r"nie\s+sprawdz[ąa]\s+si[ęe]\s+też\s+worki\s+biodegradowalne",

    # Hungarian MOHU: compostable/biodegradable bags cannot be used in the food-waste collection.
    r"komposzt[áa]lhat[óo]\s+vagy\s+leboml[óo]\s+zacsk[óo]k\s+nem\s+haszn[áa]lhat[óo]k",
]

POSITIVE_LIST_EXCLUSION_CUES = [
    r"verpakkingen\s+en\s+wegwerpproducten\s+van\s+wat\s+voor\s+materiaal\s+dan\s+ook",
    r"(?:embalagens?\s+alimentares|envases?\s+de\s+papel[- ]?cart[óo]n|envases?\s+de\s+vidrio|envases?\s+de\s+pl[áa]stico).{0,220}(?:contentor|contenedor|org[âa]nic|marr[oó]n|biorres[íi]duos)",
]

EXPLICIT_NO_PACKAGING_CUES = [
    # General packaging exclusion is only treated as relevant when it is clearly
    # inside an organics/biowaste sorting rule, not a generic recycling list.
    r"\b(organic\s*waste|organics|food\s*waste|food\s*scraps?|biowaste|bio[- ]?waste|brown\s*bin|green\s*bin|compost(?:ing)?\s*(?:bin|cart)|gft|biotonne|bac\s+brun|contenedor\s+marr[oó]n|contentor\s+castanho)\b.{0,180}\b(no|not|without|remove|separate(?:d)?\s+from|must\s+be\s+separated\s+from|do\s+not)\b.{0,100}\b(packaging|package|container|wrapping|envase|embalaje|emballage|embalagem|verpakking|forpakning|pakend|pakuot|opakowania|embala[žz]a|obaly)\b",
    r"\b(no|not|without|remove|separate(?:d)?\s+from|must\s+be\s+separated\s+from|do\s+not)\b.{0,100}\b(packaging|package|container|wrapping|envase|embalaje|emballage|embalagem|verpakking|forpakning|pakend|pakuot|opakowania|embala[žz]a|obaly)\b.{0,180}\b(organic\s*waste|organics|food\s*waste|food\s*scraps?|biowaste|bio[- ]?waste|brown\s*bin|green\s*bin|compost(?:ing)?\s*(?:bin|cart)|gft|biotonne|bac\s+brun|contenedor\s+marr[oó]n|contentor\s+castanho)\b",
    r"(no\s+packaging\s+of\s+any\s+kind|without\s+any\s+packaging|aucun\s+emballage|sans\s+emballage|sin\s+(?:ning[uú]n\s+tipo\s+de\s+)?envase|sem\s+embalagem|zonder\s+verpakking|uden\s+emballage|utan\s+förpackning|bez\s+opakowań|brez\s+embala[žz]e|bez\s+obalu)",
]


# Consolidated semantic pattern groups used by the classifiers below.
# Indicator: Acceptance Condition - official/designated item cues.
# These terms identify cases where acceptance depends on an officially supplied,
# designated, listed, or programme-approved bag/item,
# avoiding repeated long concatenations inside classifier functions.
LINER_ACCEPTANCE_PATTERNS = (
    LINER_ACCEPT_CUES
    + LINER_ACCEPTANCE_CONTEXT_CUES
    + LINER_ACCEPTANCE_MULTILINGUAL_CUES
)
BROAD_ACCEPTANCE_PATTERNS = (
    BROAD_ACCEPT_CUES
    + BROAD_ACCEPTANCE_CONTEXT_CUES
    + BROAD_ACCEPTANCE_MULTILINGUAL_CUES
)
REJECTION_PATTERNS = (
    REJECT_CUES
    + REJECTION_CONTEXT_CUES
    + REJECTION_MULTILINGUAL_CUES
)
PACKAGING_EXCLUSION_PATTERNS = (
    EXPLICIT_NO_PACKAGING_CUES
    + POSITIVE_LIST_EXCLUSION_CUES
)
REJECTION_OR_EXCLUSION_PATTERNS = REJECTION_PATTERNS + PACKAGING_EXCLUSION_PATTERNS

OFFICIAL_DESIGNATED_ITEM_CUES = [
    r"\b(official|designated|provided|supplied|distributed|issued|delivered|council[- ]?supplied|municipal[- ]?supplied|city[- ]?supplied|programme[- ]?approved|program[- ]?approved|udleverede|udleveret|h[äa]mta\s+r[äa]tt\s+p[åa]se|r[äa]tt\s+typ\s+av\s+p[åa]se|kun\s+\w{0,20}s?\s+biologisk\s+nedbrytbare\s+poser|skal\s+sorteres\s+i\s+de\s+gr[øo]nne\s+bioposer|doručen[ée]\s+vrecia\s+s[úu]\s+jedin[ée]|dodané\s+od\s+mesta|štartovacie\s+bal[íi]čky|전용봉투|종량제봉투|専用袋|指定袋|专用袋|專用袋)\b",
    r"\b(only|must|shall|exclusive|jedin[ée]|únicamente|solo|kun)\b.{0,120}\b(city|municipal|operator|council|MOVAR|TRV|mesta|municipio|ayuntamiento).{0,120}\b(bag|liner|sack|poser|bioposer|bolsas?|vrecia|vrec[úu]ška)\b",
]

ACCEPTANCE_RULES: List[Rule] = [
    Rule("No accepted route", NO_ROUTE_CUES, exclude=ACCEPT_CUES + LINER_ACCEPTANCE_PATTERNS + BROAD_ACCEPTANCE_PATTERNS, definition="No operational or compatible organics route for compostable packaging is identified.", priority=10),
    Rule("Listed items only", POSITIVE_LIST_CUES + [r"arr[êe]t[ée].{0,160}(list|annex|annexe)", r"les\s+emballages\s+et\s+d[ée]chets\s+non\s+list[ée]s"], definition="Only named/listed/approved items may enter the organics stream; generic certification is not sufficient.", priority=20),
    Rule("Controlled acceptance", CONTROLLED_CUES, exclude=REJECTION_OR_EXCLUSION_PATTERNS, definition="Acceptance exists only in controlled collection contexts such as events, closed-loop schemes, venues, or dedicated commercial collection.", priority=30),
    Rule("Liners only", LINER_ACCEPTANCE_PATTERNS, exclude=BROAD_ACCEPTANCE_PATTERNS, definition="Acceptance is limited to food-waste liners, caddy liners, collection bags, or closely equivalent collection aids.", priority=40),
    Rule("Broad acceptance", BROAD_ACCEPTANCE_PATTERNS, exclude=REJECTION_OR_EXCLUSION_PATTERNS, definition="Compostable packaging or multiple compostable item types beyond liners are accepted in the organics stream.", priority=50),
    Rule("Rejected", REJECTION_OR_EXCLUSION_PATTERNS, definition="Compostable packaging, compostable plastics, or relevant compostable items are explicitly rejected from organic waste or treated as contamination.", priority=60),
    Rule("Locally variable", FRAGMENTED_CUES, definition="Acceptance is unclear, locally variable, or fragmented across municipalities/facilities.", priority=80),
]

ACCEPTANCE_CONDITION_RULES: List[Rule] = [
    Rule("Local/facility approval required", [r"\b(local|municipal|council|collector|hauler|facility|processor|composter|operator|kommune|gemeente|commune|collectivit[ée]|azienda|impianto)\b.{0,120}\b(approval|required|accepted|permitted|check|depends|approved|zugelassen|toegestaan|autoris[ée]|approv|ammess)\b", r"\b(certification|certified|label|standard)\b.{0,120}\b(and|plus).{0,80}\b(local|facility|municipal|collector|processor)"], definition="Certification is not enough; local authority, collector, or facility approval is also required.", priority=10),
    Rule("Certification required", [r"\b(en\s*13432|ok\s*compost|seedling|keimling|kiemplant|bpi|astm\s*d\s*6400|nf\s*t\s*51[- ]?800|as\s*4736|as\s*5810|bnq|dinplus|t[üu]v\s*austria|compostabile\s*cic|certifi[ée]d?|certificad\w*|certificat\w*)\b"], definition="A named standard, certification, logo, or label is required for accepted items.", priority=20),
    Rule(
        "Official/designated item required",
        [
            r"\b(official|designated|provided|supplied|distributed|council[- ]?supplied|municipal[- ]?supplied|programme[- ]?supplied|approved|fournis?|approuv[ée]s?|distribu[ée]s?|fornit\w*|zugelassen|godkendt|hyv[äa]ksytty)\b.{0,140}\b(bag|liner|sack|item|product|sac|sachet|sacco|sacchetto|bolsa|saco|zak|zakje|pose|pussi|beutel)s?\b",
            r"(専用袋|指定袋|전용봉투|종량제봉투|专用袋|專用袋|oranžin\w*\s+maišel\w*)",
        ],
        definition="Only officially designated, supplied, or programme-approved items are accepted.",
        priority=30,
    ),
    Rule("Approved/named list required", POSITIVE_LIST_CUES, definition="Acceptance requires appearance on an approved, positive, or named-item list.", priority=40),
    Rule("Generic compostability/biodegradability requirement", [r"\b(compostable|biodegradable|composteerbaar\w*|kompostierbar\w*|compostabil\w*|biod[ée]gradable\w*|biohajoava|kompostoituva|biolagunev\w*|kompostowaln\w*|biodegradowaln\w*|可堆肥|可降解|生分解|생분해)\b"], definition="Accepted items must at least be described as compostable or biodegradable, but no named standard or approval basis is detected.", priority=90),
    Rule("Context-specific only", CONTROLLED_CUES, definition="Acceptance is limited to specific event, commercial, venue, or dedicated collection contexts.", priority=50),
]

RATIONALE_RULES: List[Rule] = [
    Rule(
        "No compatible treatment route",
        NO_ROUTE_CUES + [
            r"\bno\s+(accepted|available|suitable|compatible|clear)\s+(route|collection|treatment|facility|system)\b",
            r"\bno\s+formal\s+(organics|organic[- ]waste|biowaste|food[- ]waste)\s+(system|collection|route)\b",
        ],
        definition="No compatible collection/treatment route is identified for compostable packaging.",
        priority=10,
    ),
    Rule(
        "Explicit no-packaging / positive-list restriction",
        POSITIVE_LIST_CUES + [
            r"\b(law|regulation|ordinance|decree|directive|bylaw|statute|permitted\s+items\s+only|not\s+on\s+the\s+list|not\s+listed)\b.{0,180}\b(compostable|biodegradable|packaging|plastic|bag|liner|item|product)",
            r"\b(no|not|without|remove|separate(d)?\s+from|must\s+be\s+separated\s+from)\b.{0,90}\b(packaging|package|container|wrapping|envase|embalaje|emballage|verpakking|förpackning|forpakning|pakend)\b",
            r"\b(packaging|package|container|wrapping|envase|embalaje|emballage|verpakking|förpackning|forpakning|pakend)\b.{0,130}\b(not\s+accepted|not\s+allowed|not\s+permitted|prohibited|forbidden|exclude[d]?|remove[d]?|separate[d]?)\b",
            r"\b(sin\s+ning[uú]n\s+tipo\s+de\s+envase|sin\s+envase|separados?\s+de\s+su\s+envase|retirar\s+.*envase)\b",
            r"(verbot|vietat|interdit|prohibid|positive\s+list|non\s+list[ée]s|ne\s+sont\s+pas\s+accept[ée]s|hors\s+sacs)",
        ] + EXPLICIT_NO_PACKAGING_CUES,
        definition="Rejection or restriction follows from a legal rule, positive list, named-item rule, or explicit no-packaging rule.",
        priority=90,
    ),
    Rule(
        "AD/biogas incompatibility",
        [
            r"\b(anaerobic\s+digestion|\bAD\b|biogas|digest(?:ion|er)|m[ée]thanisation|methani[sz]ation|verg[äa]rung|biokaas|bioverg[äa]rungsanlage|biogass|biogasanlegg|biogödsel|biogødning)\b.{0,240}\b(compostable|biodegradable|bag|liner|plastic|packaging|break\s+down|clog|screen|reject|incompatib|problem)",
            r"\b(compostable|biodegradable|bag|liner|plastic|packaging)\b.{0,240}\b(anaerobic\s+digestion|\bAD\b|biogas|digest(?:ion|er)|m[ée]thanisation|methani[sz]ation|verg[äa]rung|biogass|biogas)",
        ],
        definition="The rationale specifically concerns AD/biogas compatibility.",
        priority=30,
    ),
    Rule(
        "Slow degradation / residence-time mismatch",
        [
            r"\b(residence\s*time|retention\s*time|process\s*time|degradation\s+time|decompose\s+in\s+time|break\s+down\s+in\s+time|too\s+slow|slow\s+decomposition|does\s+not\s+break\s+down|not\s+break\s+down|incomplete\s+(breakdown|degradation)|not\s+fully\s+(degrade|decompose)|6[- ]?week|12[- ]?week|longer\s+than\s+the\s+process)\b",
            r"(rottezeit|zersetzungszeit|nicht\s+sicher\s+vollst[äa]ndig\s+biologisch\s+abgebaut|langer\s+dan\s+het\s+proces|proces\s+te\s+kort|te\s+langzaam\s+af|breekt\s+te\s+langzaam|langzaam\s+afbre|biohajoava.{0,80}hidas|hitaan\s+maatumis|ikke\s+brytes\s+ned|brytes\s+ikke\s+ned|inte\s+bryts\s+ner|för\s+långsamt|nedbrytning|분해|퇴비화할|可降解塑料袋.{0,80}降解的速度比厨余垃圾慢)",
        ],
        definition="Item is rejected because it degrades more slowly than the facility process.",
        priority=40,
    ),
    Rule(
        "Pre-treatment/screening/equipment constraint",
        [
            r"\b(pre[- ]?treatment|screening|screened\s+out|front[- ]?end|sorting\s+plant|shredder|sieve|sieving|pulper|separator|machine|pump|pipe|clog|jam|remove[d]?\s+before|sorted\s+out|mechanical\s+sorting|sorting\s+facility)\b",
            r"(voorbehandeling|esik[äa]sittely|murskain|siebung|seulage|tamisage|시설|설비\s*고장|처리시설|분쇄|기계\s*부품|엉킴|纏繞後端處理設施|缠绕后端处理设施)",
        ],
        definition="Mechanical sorting, pre-treatment, screening, or equipment constraints are given.",
        priority=50,
    ),
    Rule(
        "Compost/digestate quality concern",
        [
            r"\b(compost\s+quality|digestate\s+quality|microplastic|microplastics|contamination|contaminant|foreign\s+matter|residue|pollutant|clean\s+compost|soil|agricultural\s+use|quality\s+standard|fertiliser|fertilizer|impurity|physical\s+contamination|quality\s+control|nutrient|nutritional\s+value|soil\s+amendment)\b",
            r"\b(pfas|additive|chemical|toxic|toxicity|heavy\s+metal|fluorinated|substance|contaminant\s+chemicals?|harmful\s+substances?)\b",
            r"(mikroplastik|fremdstoff|st[öo]rstoff|vervuiling|verontreiniging|schone\s+compost|kwaliteit|schadelijke\s+stoffen|compostkwaliteit|bodem|meststof|fertilisant|engrais|qualit[ée]|inquinant|퇴비|비료|품질)",
        ],
        definition="Concern about compost/digestate quality, contamination, microplastics, residues, or chemical/additive contamination.",
        priority=60,
    ),
    Rule("No compatible treatment route", [r"\b(no\s+compatible\s+(?:treatment|route)|cannot\s+be\s+(?:properly\s+)?processed|not\s+designed\s+for\s+biogas|ikke\s+produsert\s+for\s+biogassproduksjon|nie\s+mog[ąa]\s+by[ćc]\s+prawid[łl]owo\s+przetwarzane|kan\s+inte\s+.*tas\s+till\s+vara|kan\s+inte\s+tas\s+tillvara|kann\s+nicht\s+verwertet|geen\s+verwerkingsroute)\b"], definition="No compatible organic-treatment route or processing route is stated.", priority=70),
    # Lower-priority equipment-disruption cues are intentionally evaluated after
    # quality and route rationales, so a source mentioning both contamination
    # and equipment issues is counted first as a quality/contamination rationale.
    Rule("Pre-treatment/screening/equipment constraint", [r"\b(clog|jam|block|pipe|pump|agitator|equipment|machine|operational\s+problem|disruption|shredder|tangle|wrap\s+around)\b"], definition="Operational disruption or equipment interference is stated.", priority=80),
]

APPLICATION_RULES: List[Rule] = [
    Rule("Food-waste liners / collection bags", [r"\b(caddy\s*liner|kitchen\s*caddy|food[- ]?waste\s*(bag|liner)|bio\s*bag|biobag|compostable\s*(bag|liner)|green\s*bag|brown\s*bag|collection\s*(bag|liner)|organic\s*waste\s*bag|organics\s*bag|biowaste\s*bag|food\s*scraps?\s*bag)s?\b", r"\b(sac[s]?\s+(compostable|biod[ée]gradable)|sacco\s+compostabile|sacchetti\s+compostabil\w*|bolsa\s+compostable|bolsas\s+compostables|sacos?\s+compost[áa]veis|kompostierbare?\w*\s+(beutel|t[üu]ten|biobeutel)|composteerbare?\w*\s+(zak|zakken|zakjes)|biozak\w*|bioposer|matavfallsp[åa]sar|bioavfallsp[åa]sar|bioj[äa]tepussi|vre[čc]ke|worki|s[áa][čc]ky)\b", r"(生ごみ|厨余|廚餘|餐厨|음식물쓰레기|음식물류).{0,80}(袋|봉투|专用袋|專用袋|専用袋|전용봉투)"] + LINER_ACCEPTANCE_CONTEXT_CUES, exclude=[r"\b(produce\s*bag|fruit\s*(and|&|/)\s*vegetable\s*bag|shopping\s*bag|carrier\s*bag|checkout\s*bag|retail\s*bag)"], definition="Bags or liners used to collect food waste, biowaste, or kitchen/caddy organics.", priority=10),
    Rule("Food-service ware / takeaway packaging", [r"\b(food[- ]?service|serviceware|take[- ]?away|takeout|to[- ]?go|catering|event|festival|venue|closed[- ]?loop)\b", r"\b(cup|plate|bowl|tray|container|cutlery|fork|spoon|knife|straw|lid|stoviglie|posate|bicchieri|piatti|contenitori|kahvimuki|kertak[äa]ytt[öo]astia|vaisselle|gobelets?|assiettes?|couverts?|vasos?|platos?|cubiertos?)\w*\b.{0,140}\b(compostable|biodegradable|compostabil|biodegradabil|organic|umido|compost|biohajoava|kompostoituva)", r"\b(compostable|biodegradable|compostabil|biohajoava|kompostoituva)\b.{0,140}\b(cup|plate|bowl|tray|container|cutlery|fork|spoon|knife|straw|lid|serviceware|takeaway|vaisselle|gobelets?|assiettes?|couverts?)"], exclude=[r"\bcoffee\s+(filter|pod|capsule)|tea\s*bag|teabag\b"], definition="Food-service ware / takeaway packaging and serviceware such as cups, plates, bowls, trays, cutlery, containers, takeaway items, and event/venue serviceware.", priority=20),
    Rule("Tea/coffee preparation items", [r"\b(tea\s*bag|teabag|coffee\s*filter|coffee\s*(pod|capsule)|k[- ]?cup|espresso\s*capsule|teebeutel|kaffeefilter|theezakjes?|koffiefilters?|sachet[s]?\s+de\s+th[ée]|filtre[s]?\s+.*caf[ée]|capsule[s]?\s+de\s+caf[ée])s?\b"], definition="Tea bags, coffee filters, coffee pods/capsules, and related beverage-preparation items.", priority=30),
    Rule("Food-soiled paper / fibre packaging", [r"\b(food[- ]?soiled|soiled|greasy)\s+(paper|cardboard|fibre|fiber|napkin|paper\s*towel|pizza\s*box|paper\s*bag|food\s*wrapper|takeout\s*container)s?\b", r"\b(napkin|paper\s*towel|pizza\s*box|paper\s*bag|cardboard|kitchen\s*towel|keukenpapier|talouspaperi|lautasliinat|k[üu]chenpapier|paper\s+napkin|carton|serviette)s?\b.{0,140}\b(compost|organic|food waste|biowaste|biotonne|gft|bioj[äa]te|brown\s+bin)\b"], exclude=[r"\b(pdf|document|paperwork|newspaper\s+article|webpage|paper\s+published|research\s+paper)\b"], definition="Food-soiled or uncoated paper/fibre packaging and paper items, e.g. pizza boxes, napkins, paper towels, paper bags, and fibre food packaging.", priority=40),
    Rule("Shopping/produce bags", [r"\b(produce\s*bag|fruit\s*(and|&|/)\s*vegetable\s*bag|vegetable\s*bag|carrier\s*bag|shopping\s*bag|checkout\s*bag|very\s*lightweight\s*plastic\s*bag|retail\s*bag|sacchetti\s+per\s+frutta\s+e\s+verdura|shopper)s?\b"], exclude=[r"\b(food[- ]?waste|organic\s*waste|biowaste|caddy|collection)\s+(bag|liner)"], definition="Produce, carrier, shopping, checkout, or very-lightweight bags discussed as packaging rather than food-waste collection liners.", priority=50),
    Rule("Flexible films/wraps/pouches", [r"\b(mailer|postal\s*bag|e[- ]?commerce\s*packaging|garment\s*bag|poly\s*bag|protective\s*film|stretch\s*film|shrink\s*film|wrapper|wrap|pouch|pouches|wikkels|films?|folien|buste)\w*\b.{0,160}\b(compostable|biodegradable|bioplastic|compostabil|biodegradabil|organic waste|compost|brown\s+bin|green\s+cart)", r"\b(compostable|biodegradable|bioplastic|compostabil|biodegradabil)\w*\b.{0,160}\b(film|wrap|wrapper|pouch|mailer|poly\s*bag|garment\s*bag|shopping\s*bag|carrier\s*bag)"], definition="Flexible packaging such as films, wraps, wrappers, pouches, mailers, garment bags, and related flexible plastic-like packaging.", priority=60),
    Rule("Generic compostable packaging / plastics", [r"\b(compostable|biodegradable|industrially\s+compostable|certified\s+compostable|bioplastic|bio[- ]?plastic|bioplastics|composteerbaar\w*|kompostierbar\w*|compostabil\w*|biod[ée]gradable\w*|biohajoava|kompostoituva|biolagunev\w*|kompostitav\w*|biolo[šs]ko\s+razgradljiv\w*|kompostowaln\w*|biodegradowaln\w*)\b.{0,140}\b(packaging|plastic|plastics|product|products|material|materials|item|items|verpackungen|kunststoff|emballage|embalagem|imballagg|envase|verpakking|pakend|pakkauks|opakowania|embala[žz]a|obaly|muovi|pakkauks|label|sticker)\w*", r"(可堆肥|可生物降解|可降解|生分解|생분해).{0,80}(包装|包裝|塑料|塑膠|플라스틱|비닐|容器)"], definition="Generic compostable/biodegradable packaging, plastics, or products where no more specific recurring application group is detected.", priority=90),
]

APPLICATION_ORDER = [r.category for r in APPLICATION_RULES]

# ============================================================
# Loading and coding functions
# ============================================================


def load_corpus(zip_path: Path) -> List[Dict[str, Any]]:
    if not zip_path.exists():
        raise FileNotFoundError(f"Input ZIP not found: {zip_path}")
    with zipfile.ZipFile(zip_path, "r") as zf:
        candidates = [n for n in zf.namelist() if n.endswith("01_references_corpus.jsonl")]
        if not candidates:
            raise FileNotFoundError("Could not find 01_references_corpus.jsonl inside ZIP")
        rows = []
        with zf.open(candidates[0]) as f:
            for line in f:
                if line.strip():
                    rows.append(json.loads(line))
    return rows


def meta_text(row: Dict[str, Any]) -> str:
    return " | ".join([str(row.get("source_label", "")), str(row.get("source_title", "")), str(row.get("source_url", ""))])


def source_label_parts(label: str) -> Tuple[str, str]:
    parts = [p.strip() for p in str(label or "").split(" - ")]
    country = parts[0] if parts else ""
    region = parts[1] if len(parts) > 1 else ""
    return country, region


def classify_source_authority(row: Dict[str, Any]) -> RuleHit:
    """Assign institutional/geographical authority using source-label cues first.

    The regex rules remain available, but the source labels in the frozen JSONL
    corpus often contain explicit phrases such as "municipal sorting rule",
    "regional government guidance", and "waste-company guidance". These should
    take priority over broad terms like "government", which otherwise make local
    or regional sources look national.
    """
    meta = meta_text(row)
    label = norm(" | ".join([
        str(row.get("source_label", "")),
        str(row.get("source_title", "")),
        str(row.get("source_url", "")),
    ]))

    if re.search(r"\b(european union|european commission|eu framework|eur-lex|oecd|united nations|unep|supranational)\b", label):
        return RuleHit("Supranational framework", coding_basis="source_label_priority", excerpt=meta)

    # Explicit local/operational labels before broad 'government' matching.
    if re.search(r"\b(municipal sorting rule|municipal guidance|city government|city council|municipality|municipalidad|ayuntamiento|commune|gemeente|comune|borough|district council|local municipal|ville de|urząd|miasto)\b", label):
        return RuleHit("Municipal sorting rule", coding_basis="source_label_priority", excerpt=meta)
    if re.search(r"\b(waste[- ]company guidance|treatment[- ]facility|facility rule|operator|collector|processor|composting facility|biogas|public utility|epr|producer responsibility|recology|stadtreinigung|bsr|awm|ivago|saver|hvc|lipor|tratolixo|mustankorkea|biorepack|mohu|komunala)\b", label):
        return RuleHit("Waste operator or treatment-facility rule", coding_basis="source_label_priority", excerpt=meta)
    if re.search(r"\b(regional government guidance|state government|provincial|province|county|prefecture|canton|autonomous region|flanders|wallonia|catalonia|south australia|new south wales|quebec|québec|minnesota|california|ontario|alberta|niagara region|landkreis|regional rule)\b", label):
        return RuleHit("State, provincial, regional, or canton rule", coding_basis="source_label_priority", excerpt=meta)
    if re.search(r"\b(intermunicipal|metropolitan|regional waste authority|joint waste|metropole|métropole|syndicat|greater manchester|hsy)\b", label):
        return RuleHit("Intermunicipal or regional waste authority", coding_basis="source_label_priority", excerpt=meta)
    if re.search(r"\b(national|federal|ministry|ministerio|minist[eè]re|parliament|national government|country[- ]wide|bioabfv|ley\s+7/2022|boe|legifrance|gazette|federal citizen portal|national waste guidance)\b", label):
        return RuleHit("National rule or policy", coding_basis="source_label_priority", excerpt=meta)

    hit = apply_rules(meta, SOURCE_AUTHORITY_RULES, default="", source_text_for_excerpt=meta)
    if hit.category:
        return hit

    country, region = source_label_parts(str(row.get("source_label", "")))
    fallback_text = norm(" | ".join([country, region, str(row.get("source_title", "")), str(row.get("source_url", ""))]))
    if any(term in fallback_text for term in ["region", "regional", "province", "provincial", "county", "prefecture", "state", "canton", "metropolitan"]):
        return RuleHit("State, provincial, regional, or canton rule", coding_basis="fallback_metadata_inference", excerpt=meta)
    if any(term in fallback_text for term in ["municipal", "municipality", "city", "council", "commune", "gemeente", "district", "borough"]):
        return RuleHit("Municipal sorting rule", coding_basis="fallback_metadata_inference", excerpt=meta)
    if any(term in fallback_text for term in ["operator", "utility", "waste", "compost", "biogas", "services", "služby", "mpo", "olo", "voka", "snaga", "vasa"]):
        return RuleHit("Waste operator or treatment-facility rule", coding_basis="fallback_metadata_inference", excerpt=meta)
    if any(term in fallback_text for term in ["national", "federal", "ministry", "government", "parliament", "law", "regulation", "decree", "norma"]):
        return RuleHit("National rule or policy", coding_basis="fallback_metadata_inference", excerpt=meta)
    return RuleHit("Municipal sorting rule", coding_basis="fallback_metadata_inference", excerpt=meta)


def source_level_group(source_authority: str) -> str:
    if source_authority in {"National rule or policy", "Supranational framework"}:
        return "National/supranational sources"
    return "Sources below national level"


def primary_for_sensitivity(source_authority: str) -> str:
    return "Yes"


def detect_application_types(text: str) -> List[RuleHit]:
    hits = []
    txt = norm(text)
    for rule in sorted(APPLICATION_RULES, key=lambda r: r.priority):
        ex_pat, _ = first_match(rule.exclude, txt) if rule.exclude else ("", "")
        if ex_pat:
            continue
        in_pat, in_text = first_match(rule.include, txt)
        if in_pat:
            hits.append(RuleHit(rule.category, in_pat, in_text, coding_basis=rule.coding_basis, excerpt=excerpt_around(text, in_text)))
    if not hits:
        hits.append(RuleHit("Organic-waste system context only", coding_basis="fallback_no_application_match"))
    return hits


def _is_probably_negative_match(match_text: str) -> bool:
    """Return True when an apparent positive liner/bag match is actually inside
    a rejection sentence. This prevents false positives such as Dutch OVAM
    wording where compostable plastics are mentioned only to say they do not
    belong in GFT. The check is phrase-level, not citation-specific.
    """
    t = norm(match_text)
    negative = has_any([
        r"\b(niet|nooit|geen|not|never|no\s+longer|non|ne\s+sont\s+pas|ne\s+peuvent\s+pas|pas\s+accept|sin|sem|zonder|uden|utan|bez|ikke|inte|ei|nie|niso|nepatr|禁止|不可|不能|入れない|배출불가|금지)\b",
        r"\b(should\s+not|must\s+not|do\s+not|not\s+accepted|not\s+allowed|not\s+permitted|prohibited|forbidden|excluded|rejected|cannot|can\s*not|can't|can’t|why\s+can[’']?t)\b",
        r"\b(mogen\s+niet|mag\s+niet|hoort\s+niet|darf\s+nicht|d[üu]rfen\s+nicht|geh[öo]ren\s+nicht)\b",
    ], t)
    exception_or_positive = has_any([
        r"\b(hormis|sauf|except|exception|uitzondering|other\s+than|apart\s+from)\b",
        r"\b(mag\s+w[eé]l|allowed|accepted|permitted|can\s+(?:use|be\s+used|be\s+collected)|may\s+(?:use|be\s+used)|must\s+be\s+(?:paper|biodegradable|compostable))\b",
        r"\b(recomendable|obligatorio|se\s+aceptan|podem\s+ser|vanno\s+gettati|si\s+possono\s+conferire|si\s+possono\s+utilizzare\s+solo|wyj[ąa]tkiem\s+s[ąa]\s+tylko|oznaczono\s+symbolem\s+do\s+kompostowania|toegelaten|autoris[ée]s?|permitidas?)\b",
    ], t)
    return negative and not exception_or_positive


def _first_non_negative_hit(text: str, patterns: Sequence[str], default: str, *, source_text_for_excerpt: Optional[str] = None) -> RuleHit:
    txt = norm(text)
    raw = text if source_text_for_excerpt is None else source_text_for_excerpt
    for pat in patterns:
        m = compile_rx(pat).search(txt)
        if not m:
            continue
        mt = re.sub(r"\s+", " ", m.group(0)).strip()
        local_context = txt[max(0, m.start() - 120): min(len(txt), m.end() + 120)]
        # Check both the matched phrase and its immediate context, because some
        # question/FAQ wording starts with "Why can't I use..." and the positive
        # phrase itself begins at "use a compostable bag".
        if _is_probably_negative_match(mt) or _is_probably_negative_match(local_context):
            continue
        return RuleHit(default, pat, mt, coding_basis="regex_include_exclude_context_checked", excerpt=excerpt_around(raw, mt))
    return RuleHit("", coding_basis="no_non_negative_hit")


def classify_acceptance(text: str) -> RuleHit:
    """Classify acceptance using indicator-specific precedence.

    Compared with a single ordered regex list, this separates broad acceptance,
    liner exceptions, and rejection signals. This is needed because many
    operational sources say: packaging is rejected, except food-waste liners.
    Such cases should be coded as Liners only, not as simple rejection.
    """
    listed_hit = apply_rules(text, [Rule("Listed items only", POSITIVE_LIST_CUES + [r"arr[êe]t[ée].{0,160}(list|annex|annexe)", r"les\s+emballages\s+et\s+d[ée]chets\s+non\s+list[ée]s"], priority=20)], default="")

    controlled_hit = apply_rules(text, [Rule("Controlled acceptance", CONTROLLED_CUES, priority=30)], default="")
    reject_hit = apply_rules(text, [Rule("Rejected", REJECTION_OR_EXCLUSION_PATTERNS, priority=60)], default="")

    broad_hit = _first_non_negative_hit(text, BROAD_ACCEPTANCE_PATTERNS, "Broad acceptance")
    liner_hit = _first_non_negative_hit(text, LINER_ACCEPTANCE_PATTERNS, "Liners only")

    # Broad operational acceptance is rare but should override generic food-packaging
    # removal wording when the source explicitly accepts compostable packaging/items.
    if broad_hit.category and not (reject_hit.category and not has_any([r"compostable\s+packaging\s+and\s+take[- ]?away\s+items", r"stoviglie|posate|bicchieri|piatti|imballagg|cups?|plates?|bowls?|containers?|cutlery"], broad_hit.matched_text)):
        return broad_hit

    # A named-list signal that is only about caddy/food-waste bags should remain
    # liner-only; otherwise the French-style positive-list regime is retained.
    if liner_hit.category:
        return liner_hit

    if listed_hit.category:
        return listed_hit

    no_route_hit = apply_rules(text, [Rule("No accepted route", NO_ROUTE_CUES, priority=10)], default="")
    if no_route_hit.category and not broad_hit.category and not liner_hit.category:
        return no_route_hit

    if controlled_hit.category and not reject_hit.category:
        return controlled_hit

    if reject_hit.category:
        return reject_hit

    local_hit = apply_rules(text, [Rule("Locally variable", FRAGMENTED_CUES, priority=80)], default="")
    if local_hit.category:
        return local_hit

    return RuleHit("No explicit rule", coding_basis="fallback_default")


def classify_treatment_route(text: str) -> RuleHit:
    return apply_rules(text, TREATMENT_ROUTE_RULES, default="Treatment route not stated")


def treatment_route_coding_note(hit: RuleHit) -> str:
    """Explain treatment-route not-stated cases without adding a new treatment route category.

    Many local rules specify only the collection/bin route (brown bin, green cart,
    GFT, FOGO, biotonne, etc.) and do not state whether material goes to
    composting, AD/biogas, or another recovery route. These remain coded as
    Treatment route not stated, but the note separates collection-only evidence
    from true fallback cases.
    """
    if hit.category != "Treatment route not stated":
        return "Named downstream treatment route detected"
    if hit.matched_rule:
        return "Collection/bin route stated only; downstream treatment technology not stated"
    return "No collection-route or downstream-treatment cue detected"


def classify_certification_basis(text: str) -> RuleHit:
    return apply_rules(text, CERTIFICATION_RULES, default="No standard stated")


def classify_acceptance_condition(text: str, acceptance_category: str, cert_category: str = "") -> RuleHit:
    if acceptance_category in {"Rejected", "No accepted route", "No explicit rule", "Locally variable"}:
        return RuleHit("Not applicable - not accepted/no explicit route", coding_basis="derived_from_acceptance")

    hit = apply_rules(text, ACCEPTANCE_CONDITION_RULES, default="No additional condition stated")
    if hit.category != "No additional condition stated":
        return hit

    # Secondary official/designated-item inference is intentionally applied only
    # after direct condition rules. This preserves the modelling distinction that
    # explicit standards or context restrictions should not be overwritten by a
    # broader official-bag phrase elsewhere on the page.
    official_hit = apply_rules(
        text,
        [Rule("Official/designated item required", OFFICIAL_DESIGNATED_ITEM_CUES, priority=1)],
        default="",
    )
    if official_hit.category:
        return official_hit

    if cert_category in {"EN 13432 / OK compost / Seedling", "BPI / ASTM / CMA", "OK compost HOME / NF T 51-800 / AS 5810", "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)"}:
        return RuleHit("Certification required", coding_basis="derived_from_certification_basis")
    if cert_category == "Government/programme-approved":
        return RuleHit("Official/designated item required", coding_basis="derived_from_certification_basis")
    if cert_category == "Generic compostable/biodegradable claim only":
        return RuleHit("Generic compostability/biodegradability requirement", coding_basis="derived_from_certification_basis")
    return hit


def classify_rejection_rationale(text: str, acceptance_category: str) -> RuleHit:
    if acceptance_category not in {"Rejected", "No accepted route", "Liners only", "Listed items only", "Controlled acceptance"}:
        return RuleHit("Not applicable", coding_basis="derived_from_acceptance")
    hit = apply_rules(text, RATIONALE_RULES, default="Reason not specified")
    return hit


def decide_application(application_type: str, acceptance_category: str, text: str, acceptance_condition: str = "", certification_basis: str = "") -> Tuple[str, str]:
    """Four-category application decision before binary scenario conversion.

    Important modelling distinction
    -------------------------------
    Certification or generic compostability is treated as product qualification,
    not as an operational restriction. Therefore, EN 13432/BPI/OK compost etc.
    do not by themselves downgrade an application from Accepted to Accepted
    with conditions.

    "Accepted with conditions" is reserved for restrictions that change the
    route/product context: official/designated items only, local/facility
    approval, positive-list inclusion, event/commercial/dedicated collection,
    or another controlled route beyond ordinary household organics collection.
    """
    if application_type == "Organic-waste system context only":
        return "Unclear or not stated", "No product/application-specific item detected."

    operational_restriction = acceptance_condition in {
        "Official/designated item required",
        "Local/facility approval required",
        "Approved/named list required",
        "Context-specific only",
    } or certification_basis == "Government/programme-approved"

    certification_only = acceptance_condition in {
        "Certification required",
        "Generic compostability/biodegradability requirement",
        "No additional condition stated",
        "No condition / not applicable",
        "",
    } and certification_basis not in {"Government/programme-approved"}

    if acceptance_category == "Broad acceptance":
        if operational_restriction:
            return "Accepted with conditions", "Application is accepted, but only under an operational/product restriction such as official approval, local/facility approval, named-list inclusion, or a controlled route."
        return "Accepted", "Application is accepted in the ordinary organics route; certification or generic compostability is treated as product qualification, not a scenario restriction."

    if acceptance_category in {"Rejected", "No accepted route"}:
        return "Rejected", "Source-level evidence rejects compostable packaging/items or identifies no accepted organics route."

    if acceptance_category == "Liners only":
        if application_type == "Food-waste liners / collection bags":
            if operational_restriction:
                return "Accepted with conditions", "Food-waste liners are accepted only as official/designated, locally/facility-approved, or otherwise restricted collection items."
            return "Accepted", "Food-waste liners/collection bags are accepted in the ordinary organics route; certification or compostability wording is treated as product qualification."
        return "Rejected", "Source-level rule is liner-only; this broader application category is outside the accepted application scope."

    if acceptance_category == "Listed items only":
        positive_list_applications = {
            "Food-waste liners / collection bags",
            "Tea/coffee preparation items",
            "Food-soiled paper / fibre packaging",
            "Shopping/produce bags",
        }
        if application_type in positive_list_applications:
            return "Accepted with conditions", "Application is accepted only if it is one of the listed/approved item types under a positive-list or named-item rule."
        return "Rejected", "Positive-list/named-item regime excludes generic or non-listed compostable packaging applications."

    if acceptance_category == "Controlled acceptance":
        controlled_applications = {
            "Food-service ware / takeaway packaging",
            "Food-waste liners / collection bags",
            "Generic compostable packaging / plastics",
            "Food-soiled paper / fibre packaging",
        }
        if application_type in controlled_applications:
            return "Accepted with conditions", "Application is accepted only in a controlled, facility-approved, event, commercial, or dedicated collection context."
        return "Unclear or not stated", "Controlled-route evidence does not clearly cover this application category."

    return "Unclear or not stated", "Application is mentioned, but no explicit application-level acceptance or rejection decision is stated."

def binary_scenario_row(app_type: str, counts: Dict[str, int]) -> Dict[str, Any]:
    a = counts.get("Accepted", 0)
    c = counts.get("Accepted with conditions", 0)
    r = counts.get("Rejected", 0)
    u = counts.get("Unclear or not stated", 0)
    lower_den = a + c + r
    central_den = a + c + r
    upper_den = a + c + r + u
    return {
        "application_type": app_type,
        "n_accepted": a,
        "n_accepted_with_conditions": c,
        "n_rejected": r,
        "n_unclear_or_not_stated": u,
        "evidence_records": a + c + r + u,
        "lower_bound_denominator": lower_den,
        "lower_bound_compatible_share": a / lower_den if lower_den else 0,
        "central_denominator": central_den,
        "central_compatible_share": (a + c) / central_den if central_den else 0,
        "upper_bound_denominator": upper_den,
        "upper_bound_compatible_share": (a + c + u) / upper_den if upper_den else 0,
    }


# ============================================================
# Analysis route 3: certification sufficiency
# ============================================================

NO_CERTIFICATION_BASIS_RAW = "No standard stated"
NO_CERTIFICATION_BASIS_DISPLAY = "No named standard/approval stated"

# Certification Sufficiency is intentionally evaluated only for records whose
# source-level Certification and Approval Basis is a certification, standard, or
# official approval basis. Generic compostable/biodegradable wording and sources
# with no named standard/approval are retained in Certification and Approval
# Basis, but they are excluded from the Certification Sufficiency denominator.
CERTIFICATION_RELEVANT_BASIS = {
    "EN 13432 / OK compost / Seedling",
    "Government/programme-approved",
    "Government/programme approval",  # display-label compatibility
    "BPI / ASTM / CMA",
    "OK compost HOME / NF T 51-800 / AS 5810",
    "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
    # legacy labels kept so older intermediate CSVs remain readable
    "BNQ / AS 4736 / DINplus / national standards",
    "National standards (BNQ / AS 4736 / DINplus etc.)",
    "National standards",
}

CERTIFICATION_EXCLUDED_BASIS = {
    "Generic compostable/biodegradable claim only",
    "Generic compostable/biodegradable wording only",
    NO_CERTIFICATION_BASIS_RAW,
    NO_CERTIFICATION_BASIS_DISPLAY,
    "",
}

CERTIFICATION_SUFFICIENCY_ORDER = [
    "Certification accepted",
    "Certification conditionally accepted",
    "Certification rejected",
    "Certification unclear or not stated",
]


def is_certification_relevant_basis(certification_basis: Any) -> bool:
    """Return True for certification, standard, or official approval bases.

    Generic compostable/biodegradable wording is not treated as certification.
    """
    if pd.isna(certification_basis):
        return False
    value = str(certification_basis).strip()
    return value in CERTIFICATION_RELEVANT_BASIS


def certification_basis_type(certification_basis: Any) -> str:
    """Classify source-level Certification and Approval Basis for Route 3."""
    if is_certification_relevant_basis(certification_basis):
        return "certification_relevant"
    if pd.isna(certification_basis):
        return "excluded_no_named_standard_or_approval"
    value = str(certification_basis).strip()
    if value in {"Generic compostable/biodegradable claim only", "Generic compostable/biodegradable wording only"}:
        return "excluded_generic_compostability_wording"
    return "excluded_no_named_standard_or_approval"


def has_named_certification(certification_basis: Any) -> bool:
    """Return True only for certification, standard, or official approval bases."""
    return is_certification_relevant_basis(certification_basis)


def certification_sufficiency_status(row: pd.Series) -> Tuple[str, str, int]:
    """Classify Certification Sufficiency for certification-relevant records.

    This indicator is derived only after filtering to source × application-group
    records whose source-level Certification and Approval Basis is a certification,
    standard, or official approval category. It then maps the application-level
    Application Decision to four mutually exclusive categories.

    Score is ordinal only for display/ordering:
    0 = Certification accepted;
    1 = Certification conditionally accepted;
    2 = Certification rejected;
    3 = Certification unclear or not stated.
    """
    decision = str(row.get("application_decision", ""))

    if decision == "Accepted":
        return (
            "Certification accepted",
            "The source mentions a certification, standard, or official approval basis, and this application group is accepted in the ordinary organic-waste route.",
            0,
        )

    if decision == "Accepted with conditions":
        return (
            "Certification conditionally accepted",
            "The source mentions a certification, standard, or official approval basis, but acceptance still depends on additional operational gatekeeping such as local/facility approval, official/designated items, positive-list inclusion, programme approval, event/commercial collection, or another controlled route.",
            1,
        )

    if decision == "Rejected":
        return (
            "Certification rejected",
            "The source mentions a certification, standard, or official approval basis, but this application-group record is rejected from the organic-waste route.",
            2,
        )

    return (
        "Certification unclear or not stated",
        "The source mentions a certification, standard, or official approval basis, but the application-level decision is unclear or not stated.",
        3,
    )


def ordered_category_table(df: pd.DataFrame, index_col: str, category_col: str, order: Sequence[str], index_order: Optional[Sequence[str]] = None) -> pd.DataFrame:
    """Create a wide count/share table using fixed category and optional row order."""
    if df.empty:
        if index_order is None:
            return pd.DataFrame()
        out = pd.DataFrame({index_col: list(index_order)})
        for cat in order:
            out[cat] = 0
        out["evidence_records"] = 0
        for cat in order:
            out[f"share__{cat}"] = pd.NA
        return out

    counts = pd.crosstab(df[index_col], df[category_col]).reset_index()
    if index_order is not None:
        counts = pd.DataFrame({index_col: list(index_order)}).merge(counts, on=index_col, how="left")
    for cat in order:
        if cat not in counts.columns:
            counts[cat] = 0
    count_cols = [cat for cat in order if cat in counts.columns]
    counts[count_cols] = counts[count_cols].fillna(0).astype(int)
    counts = counts[[index_col] + count_cols]
    counts["evidence_records"] = counts[count_cols].sum(axis=1)
    for cat in count_cols:
        counts[f"share__{cat}"] = counts[cat] / counts["evidence_records"].replace({0: pd.NA})
    return counts


def certification_sufficiency_bubble_matrix(by_app: pd.DataFrame) -> pd.DataFrame:
    """Return long-form matrix data for the bubble chart.

    One row = Application Group × Certification Sufficiency category. This is
    the clean base table for plotting a bubble matrix.
    """
    if by_app is None or by_app.empty:
        rows = []
        for app_type in APPLICATION_ORDER:
            for status in CERTIFICATION_SUFFICIENCY_ORDER:
                rows.append({
                    "application_type": app_type,
                    "certification_sufficiency_status": status,
                    "n": 0,
                    "application_denominator": 0,
                    "share_within_application": pd.NA,
                })
        return pd.DataFrame(rows)

    rows = []
    for _, r in by_app.iterrows():
        app_type = r.get("application_type", "")
        denom = int(r.get("evidence_records", 0) or 0)
        for status in CERTIFICATION_SUFFICIENCY_ORDER:
            n = int(r.get(status, 0) or 0)
            rows.append({
                "application_type": app_type,
                "certification_sufficiency_status": status,
                "n": n,
                "application_denominator": denom,
                "share_within_application": n / denom if denom else pd.NA,
            })
    return pd.DataFrame(rows)


def build_certification_sufficiency_tables(ref_df: pd.DataFrame, app_df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Build certification-sufficiency outputs from certification-relevant records.

    Denominator: source × application-group records whose source-level
    Certification and Approval Basis is a certification, standard, or official
    approval basis. Generic compostable/biodegradable wording and no-standard
    sources are outside this denominator.
    """
    empty = pd.DataFrame()
    if app_df.empty:
        return {
            "certification_sufficiency_long_display": empty,
            "certification_sufficiency_summary_display": empty,
            "certification_sufficiency_by_application_display": empty,
            "certification_sufficiency_bubble_matrix_display": empty,
            "certification_sufficiency_by_certification_basis_display": empty,
            "certification_sufficiency_by_source_authority_display": empty,
        }

    ref_cols = [
        "citation_id",
        "country_or_region",
        "Source Authority",
        "Stated Organic-waste Treatment Route",
        "Acceptance",
        "Acceptance Condition",
        "Certification or Approval Basis",
        "Rejection Rationale",
    ]
    merged_all = app_df.merge(ref_df[[c for c in ref_cols if c in ref_df.columns]], on=["citation_id", "country_or_region"], how="left")
    merged_all["certification_basis_type"] = merged_all["Certification or Approval Basis"].map(certification_basis_type)
    merged_all["certification_relevant_record"] = merged_all["Certification or Approval Basis"].map(lambda x: "Yes" if is_certification_relevant_basis(x) else "No")

    merged = merged_all[merged_all["certification_relevant_record"].eq("Yes")].copy()

    if merged.empty:
        by_app = ordered_category_table(merged, "application_type", "certification_sufficiency_status", CERTIFICATION_SUFFICIENCY_ORDER, index_order=APPLICATION_ORDER)
        return {
            "certification_sufficiency_long_display": merged,
            "certification_sufficiency_summary_display": pd.DataFrame({
                "certification_sufficiency_status": CERTIFICATION_SUFFICIENCY_ORDER,
                "n": [0] * len(CERTIFICATION_SUFFICIENCY_ORDER),
                "denominator": [0] * len(CERTIFICATION_SUFFICIENCY_ORDER),
                "share": [pd.NA] * len(CERTIFICATION_SUFFICIENCY_ORDER),
                "share_percent": [pd.NA] * len(CERTIFICATION_SUFFICIENCY_ORDER),
                "denominator_basis": ["Certification-relevant source × application-group records."] * len(CERTIFICATION_SUFFICIENCY_ORDER),
            }),
            "certification_sufficiency_by_application_display": by_app,
            "certification_sufficiency_bubble_matrix_display": certification_sufficiency_bubble_matrix(by_app),
            "certification_sufficiency_by_certification_basis_display": empty,
            "certification_sufficiency_by_source_authority_display": empty,
        }

    status_basis_score = merged.apply(certification_sufficiency_status, axis=1, result_type="expand")
    merged["certification_sufficiency_status"] = status_basis_score[0]
    merged["certification_sufficiency_basis"] = status_basis_score[1]
    merged["certification_sufficiency_score"] = status_basis_score[2].astype("Int64")
    merged["has_named_certification_or_approval"] = "Yes"

    status_counts = merged["certification_sufficiency_status"].value_counts().reindex(CERTIFICATION_SUFFICIENCY_ORDER, fill_value=0)
    denom = int(status_counts.sum())
    summary = pd.DataFrame({
        "certification_sufficiency_status": status_counts.index,
        "n": status_counts.values.astype(int),
        "denominator": denom,
    })
    summary["share"] = summary["n"] / summary["denominator"].replace({0: pd.NA})
    summary["share_percent"] = 100 * summary["share"]
    summary["denominator_basis"] = "Certification-relevant source × application-group records: records whose source-level Certification and Approval Basis is a certification, standard, or official approval category. Generic compostable/biodegradable wording and no named standard/approval are excluded."

    by_app = ordered_category_table(merged, "application_type", "certification_sufficiency_status", CERTIFICATION_SUFFICIENCY_ORDER, index_order=APPLICATION_ORDER)
    if not by_app.empty:
        means = merged.groupby("application_type")["certification_sufficiency_score"].mean()
        by_app["mean_certification_sufficiency_score"] = by_app["application_type"].map(means)
        by_app["denominator_basis"] = "Certification-relevant source × application-group records for the same application group."

    bubble_matrix = certification_sufficiency_bubble_matrix(by_app)
    bubble_matrix["denominator_basis"] = "Certification-relevant source × application-group records for the same application group; this long table is the direct base data for the certification-sufficiency bubble chart."

    by_cert = ordered_category_table(merged, "Certification or Approval Basis", "application_decision", ["Accepted", "Accepted with conditions", "Rejected", "Unclear or not stated"])
    if not by_cert.empty:
        by_cert["denominator_basis"] = "Certification-relevant source × application-group records for the same Certification and Approval Basis category."
    by_source = ordered_category_table(merged, "Source Authority", "certification_sufficiency_status", CERTIFICATION_SUFFICIENCY_ORDER)
    if not by_source.empty:
        by_source["denominator_basis"] = "Certification-relevant source × application-group records for the same Source Authority category."

    detail_cols = [
        "citation_id",
        "country_or_region",
        "source_label",
        "source_title",
        "application_type",
        "application_decision",
        "Certification or Approval Basis",
        "Source Authority",
        "Stated Organic-waste Treatment Route",
        "Acceptance",
        "Acceptance Condition",
        "Rejection Rationale",
        "certification_sufficiency_status",
        "certification_sufficiency_score",
        "certification_sufficiency_basis",
        "certification_basis_type",
        "certification_relevant_record",
        "has_named_certification_or_approval",
        "decision_basis",
        "source_url",
    ]
    detail = merged[[c for c in detail_cols if c in merged.columns]].copy()

    return {
        "certification_sufficiency_long_display": detail,
        "certification_sufficiency_summary_display": summary,
        "certification_sufficiency_by_application_display": by_app,
        "certification_sufficiency_bubble_matrix_display": bubble_matrix,
        "certification_sufficiency_by_certification_basis_display": by_cert,
        "certification_sufficiency_by_source_authority_display": by_source,
    }


def share_table(df: pd.DataFrame, indicator: str, col: str, question: str, denominator_basis: str, mask: Optional[pd.Series] = None) -> pd.DataFrame:
    sdf = df.loc[mask].copy() if mask is not None else df.copy()
    denom = len(sdf)
    vc = sdf[col].fillna("Not coded").value_counts(dropna=False)
    rows = []
    for cat, n in vc.items():
        rows.append({
            "indicator": indicator,
            "question_answered": question,
            "category": cat,
            "n": int(n),
            "denominator": int(denom),
            "share": n / denom if denom else 0,
            "share_percent": 100 * n / denom if denom else 0,
            "denominator_basis": denominator_basis,
        })
    return pd.DataFrame(rows).sort_values(["indicator", "n", "category"], ascending=[True, False, True])


def treatment_route_not_stated_audit(ref_df: pd.DataFrame) -> pd.DataFrame:
    """Audit why treatment-route-not-stated cases remain in the backup category."""
    needed = [
        "citation_id", "country_or_region", "subnational_or_context", "source_label",
        "Source Authority", "Stated Organic-waste Treatment Route",
        "Treatment Route Coding Note", "Acceptance", "source_url",
    ]
    if ref_df.empty or "Stated Organic-waste Treatment Route" not in ref_df.columns:
        return pd.DataFrame(columns=needed)
    mask = ref_df["Stated Organic-waste Treatment Route"].eq("Treatment route not stated")
    return ref_df.loc[mask, [c for c in needed if c in ref_df.columns]].copy()


CODEBOOK_INCLUDED_EXAMPLES = {
    "Composting": "composting, industrial composting, in-vessel composting, compost site",
    "AD / biogas": "anaerobic digestion, AD, biogas, biomethane, methanisation, vergärung",
    "AD + composting": "AD followed by composting; biogas plus compost; digestion and composting both stated",
    "Other valorisation": "animal feed, insect treatment, rendering, food-waste recovery route other than composting/AD",
    "Treatment route not stated": "the source gives no downstream treatment technology, even if it may state a collection/bin route",
    "Broad acceptance": "certified compostable packaging/serviceware allowed in green/brown/organics bin beyond liners",
    "Liners only": "compostable caddy liners, food-waste bags, official collection bags only",
    "Listed items only": "positive list, named item list, approved products list, ministerial annex/list items",
    "Controlled acceptance": "closed-loop event, dedicated commercial collection, venue-specific serviceware collection",
    "Rejected": "do not put compostable plastics/packaging in organics; treated as contamination/residual waste",
    "No explicit rule": "organics source is relevant but gives no direct compostable-packaging rule",
    "Locally variable": "depends on municipality/local collector; unclear; locally variable",
    "No accepted route": "no compatible route for compostable packaging despite organic-waste context",
    "Certification required": "EN 13432 required; BPI certified; OK compost required",
    "Local/facility approval required": "certification plus local authority, collector, facility, or approved-list requirement",
    "Official/designated item required": "council-supplied bag, government-designated bag, official printed/grid bag",
    "Approved/named list required": "positive list, approved products list, named annex/list items",
    "Context-specific only": "event-only, commercial route only, dedicated collection only",
    "No condition / not applicable": "accepted route found, but no additional condition is stated",
    "OK compost HOME / NF T 51-800 / AS 5810": "OK compost HOME, NF T 51-800, AS 5810, home-compostable mark",
    "EN 13432 / OK compost / Seedling": "EN 13432, OK compost, Seedling/Keimling/Kiemplant, DIN CERTCO, Compostabile CIC",
    "BPI / ASTM / CMA": "BPI, ASTM D6400, ASTM D6868, Compost Manufacturing Alliance",
    "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)": "BNQ, AS 4736, DINplus, JIS, or other country-specific compostability standards/certifications",
    "Government/programme-approved": "official/provided/approved bag or programme-approved product",
    "Generic compostable/biodegradable claim only": "compostable, biodegradable, biohajoava, kompostierbar without named standard",
    "No standard stated": "no EN/BPI/OK compost/home-compostability or approval basis detected",
    "Slow degradation / residence-time mismatch": "does not break down in time; too slow; short residence time; incomplete degradation",
    "Explicit no-packaging / positive-list restriction": "explicit no-packaging rule, positive-list restriction, or named-item/legal exclusion",
    "Pre-treatment/screening/equipment constraint": "screened out, pre-treatment removes bags/plastics, sorting equipment removes films",
    "AD/biogas incompatibility": "biogas/AD process incompatible; clogs pipes; not designed for anaerobic digestion",
    "Compost/digestate quality concern": "compost/digestate quality, soil contamination, visual contamination, agricultural use concern",
    "Chemical/additive concern": "PFAS, additives, chemicals, toxins, coating, contamination concern",
    "No compatible treatment route": "no facility/infrastructure/service able to handle compostable packaging",
    "Pre-treatment/screening/equipment constraint": "clogging, jamming, shredders, equipment disruption, operational problems",
    "Reason not specified": "rejection detected but no specific mechanism stated",
    "Not applicable": "source does not reject or restrict compostable packaging/items",
    "Food-waste liners / collection bags": "caddy liner, food-waste bag, biobag, collection liner, biojätepussi",
    "Food-service ware / takeaway packaging": "compostable cups, plates, bowls, trays, cutlery, food-service containers",
    "Tea/coffee preparation items": "tea bags, coffee filters, coffee pods/capsules",
    "Food-soiled paper / fibre packaging": "pizza boxes, napkins, paper towels, soiled paper/cardboard, fibre food packaging",
    "Shopping/produce bags": "produce bags, fruit/vegetable bags, carrier/shopping/checkout bags",
    "Flexible films/wraps/pouches": "compostable film, wrap, pouch, mailer, garment bag, protective film",
    "Generic compostable packaging / plastics": "generic compostable packaging/products/plastics without a more specific recurring application group",
    "Accepted": "application is accepted in the ordinary organics route; named certification or generic compostability is product qualification, not a scenario restriction",
    "Accepted with conditions": "application is accepted only under an operational/product restriction: official/designated item, local/facility approval, positive-list inclusion, event/commercial/dedicated route, or other controlled route",
    "Rejected": "application is rejected, treated as contamination/residual waste, has no accepted route, or is outside a liner/listed-item rule",
    "Unclear or not stated": "application is mentioned but no explicit application-level acceptance/rejection decision can be assigned",
    "Lower-bound compatible share": "only clearly accepted application decisions count as compatible",
    "Central compatible share": "accepted and accepted-with-conditions count as compatible",
    "Upper-bound compatible share": "accepted, conditional, and unclear/not-stated count as compatible; rejected remains incompatible",
    "National/supranational sources": "national rule or policy; supranational framework",
    "Sources below national level": "municipal, operator/facility, state/regional/canton, intermunicipal/regional waste authority",
    "Certification accepted": "certification-relevant basis + Application Decision = Accepted",
    "Certification conditionally accepted": "certification-relevant basis + Application Decision = Accepted with conditions",
    "Certification rejected": "certification-relevant basis + Application Decision = Rejected",
    "Certification unclear or not stated": "certification-relevant basis + Application Decision = Unclear or not stated",
}

CODEBOOK_EXCLUDED_EXAMPLES = {
    "Composting": "AD-only/biogas-only sources",
    "AD / biogas": "composting-only sources",
    "AD + composting": "sources mentioning only one treatment route",
    "Other valorisation": "ordinary composting or AD routes",
    "Route unclear": "sources explicitly naming composting, AD, biogas, or another route",
    "Broad acceptance": "liner-only, positive-list-only, event-only, or unclear rules",
    "Liners only": "general food-service packaging acceptance; broad acceptance",
    "Listed items only": "generic certification acceptance without a named list/item constraint",
    "Controlled acceptance": "ordinary household brown-bin acceptance",
    "Rejected": "conditional or positive-list acceptance cases",
    "No explicit rule": "explicit accept/reject/condition language",
    "Locally variable": "clear national/municipal/facility rule",
    "No accepted route": "sources with a stated compatible route for some compostable packaging",
    "Certification required": "generic compostable wording without named certification/standard",
    "Local/facility approval required": "certification alone with no local/facility approval cue",
    "Official/designated item required": "generic certified bags sold at retail without official designation",
    "Approved/named list required": "broad acceptance without list/annex/named item restriction",
    "Context-specific only": "ordinary household acceptance not limited to event/commercial/dedicated context",
    "No condition / not applicable": "explicit certification, approval, official item, or context-specific condition",
    "OK compost HOME / NF T 51-800 / AS 5810": "industrial-only EN 13432/OK compost industrial unless home standard also present",
    "EN 13432 / OK compost / Seedling": "home-only standard or North American BPI/ASTM-only evidence",
    "BPI / ASTM / CMA": "European EN/OK compost-only evidence without BPI/ASTM/CMA",
    "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)": "generic industrial standard with no country-specific standard/certification cue",
    "Government/programme-approved": "generic compostable wording with no official/designated/programme approval cue",
    "Generic compostable/biodegradable claim only": "named standard, certification, or official approval basis",
    "No standard stated": "any named standard, mark, certification, or approval basis",
    "Slow degradation / residence-time mismatch": "rejection based only on law/list, not degradation time",
    "Explicit no-packaging / positive-list restriction": "facility-process rejection without legal/list/no-packaging cue",
    "Pre-treatment/screening/equipment constraint": "general degradation/slowness with no screening/pre-treatment cue",
    "AD/biogas incompatibility": "ordinary composting residence-time problem",
    "Compost/digestate quality concern": "equipment disruption without quality/digestate/soil concern",
    "Chemical/additive concern": "physical contamination only with no chemical/additive cue",
    "No compatible treatment route": "specific route exists for accepted compostable items",
    "Pre-treatment/screening/equipment constraint": "legal exclusion with no equipment/process-disruption wording",
    "Reason not specified": "a specific rationale cue is detected",
    "Not applicable": "source is coded as rejected/restricted with a rationale",
    "Food-waste liners / collection bags": "carrier/produce bags unless linked to food-waste collection; general packaging",
    "Food-service ware / takeaway packaging": "tea/coffee preparation items, liners, generic packaging without food-service item cue",
    "Tea/coffee preparation items": "generic cups or containers not used for tea/coffee preparation",
    "Food-soiled paper / fibre packaging": "research papers, source documents, generic paper/cardboard recycling instructions",
    "Shopping/produce bags": "food-waste caddy liners or official collection bags",
    "Flexible films/wraps/pouches": "paper bags, rigid containers, generic packaging without film/wrap/pouch cue",
    "Generic compostable packaging / plastics": "more specific recurring application group already detected",
    "Accepted": "operationally restricted acceptance, rejected, or unclear application decisions",
    "Accepted with conditions": "ordinary acceptance based only on certification/product qualification, outright rejection, or unclear evidence",
    "Rejected": "accepted/conditional applications or unclear cases",
    "Unclear or not stated": "explicit accepted, conditional, or rejected applications",
    "Lower-bound compatible share": "conditional/unclear decisions are not counted compatible in this scenario",
    "Central compatible share": "unclear/not stated is excluded from denominator",
    "Upper-bound compatible share": "rejected decisions are never counted compatible",
    "National/supranational sources": "municipal/operator/regional operational sources",
    "Sources below national level": "national/supranational sources",
    "Certification accepted": "certification-relevant records with conditional, rejected, or unclear application decisions; generic/no-standard records",
    "Certification conditionally accepted": "certification-relevant records with accepted, rejected, or unclear application decisions; generic/no-standard records",
    "Certification rejected": "certification-relevant records with accepted, conditionally accepted, or unclear application decisions; generic/no-standard records",
    "Certification unclear or not stated": "certification-relevant records with accepted, conditionally accepted, or rejected application decisions; generic/no-standard records",
}

def codebook_rows() -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    def include_regex_examples(rule: Rule) -> str:
        if rule.include:
            return " | ".join(rule.include[:4])
        return "No direct include regex: category is assigned by fallback/default or derived mapping."

    def exclude_regex_examples(rule: Rule) -> str:
        if rule.exclude:
            return " | ".join(rule.exclude[:3])
        return "No category-specific exclude regex."

    def combined_regex_examples(rule: Rule) -> str:
        return f"INCLUDE: {include_regex_examples(rule)} ; EXCLUDE: {exclude_regex_examples(rule)}"

    def add_from_rules(section: str, indicator: str, question: str, rules: Sequence[Rule]):
        seen_categories = set()
        for i, r in enumerate(sorted(rules, key=lambda x: x.priority), 1):
            if r.category in seen_categories:
                continue
            seen_categories.add(r.category)
            rows.append({
                "section": section,
                "indicator": indicator,
                "question_answered": question,
                "category": r.category,
                "definition": r.definition,
                "included_examples": r.included_examples or CODEBOOK_INCLUDED_EXAMPLES.get(r.category, "See regex/rule examples for representative matched wording."),
                "excluded_examples": r.excluded_examples or CODEBOOK_EXCLUDED_EXAMPLES.get(r.category, "Cases outside the category definition are excluded; see regex/rule examples."),
                "include_regex_examples": include_regex_examples(r),
                "exclude_regex_examples": exclude_regex_examples(r),
                "regex_or_rule_examples": combined_regex_examples(r),
                "coding_basis": r.coding_basis,
                "priority_or_mapping_note": r.note or f"Priority {i}; exclusions are evaluated before inclusion for this category.",
            })

    add_from_rules("Evidence-base descriptor", "Source Authority", "What level/type of source provides the evidence?", SOURCE_AUTHORITY_RULES)
    add_from_rules("Source-level indicator", "Stated Organic-waste Treatment Route", "What organic-waste treatment route is stated?", TREATMENT_ROUTE_RULES + [Rule("Treatment route not stated", [], definition="The source does not state a downstream treatment technology such as composting, AD/biogas, or other organic recovery.", coding_basis="fallback_default")])
    add_from_rules("Source-level indicator", "Acceptance", "How are compostable packaging/items treated in organic-waste systems?", ACCEPTANCE_RULES + [Rule("No explicit rule", [], definition="The source discusses organics/biowaste but does not explicitly address compostable packaging acceptance or rejection.", coding_basis="fallback_default")])
    add_from_rules("Intermediate decision input", "Acceptance Condition", "Which gatekeeping condition is attached to acceptance before application and certification-sufficiency mapping?", ACCEPTANCE_CONDITION_RULES + [Rule("No condition / not applicable", [], definition="No condition is specified, or the item/rule is not accepted. This indicator is retained as an intermediate input, not as a standalone headline result, because it overlaps conceptually with the derived certification-sufficiency indicator.", coding_basis="derived/fallback")])
    add_from_rules("Source-level indicator", "Certification and Approval Basis", "What standard, certification, label, or approval basis is mentioned as the upstream product-qualification or approval signal?", CERTIFICATION_RULES + [Rule("No standard stated", [], definition="No named standard, certification, label, or programme/facility approval basis is detected.", coding_basis="fallback_default")])
    add_from_rules("Source-level indicator", "Rejection Rationale", "Why are compostable packaging/items rejected or restricted?", RATIONALE_RULES + [Rule("Reason not specified", [], definition="The source rejects or restricts items, but no specific rationale is detected.", coding_basis="fallback_default"), Rule("Not applicable", [], definition="The source is not coded as rejecting or restricting compostable packaging/items.", coding_basis="derived_from_acceptance")])
    add_from_rules("Application-level indicator", "Application Group", "Which product/application group does the evidence refer to?", APPLICATION_RULES)
    app_decision_definitions = {
        "Accepted": "The application is accepted in the ordinary organics route. A named certification, label, or generic compostability requirement may still apply, but this is treated as product qualification rather than a scenario-level operational restriction.",
        "Accepted with conditions": "The application is accepted only under an operational/product restriction beyond ordinary certification: official/designated item, local/facility approval, positive-list inclusion, event/commercial/dedicated route, or another controlled-route condition.",
        "Rejected": "The application is explicitly rejected from the organics route, treated as contamination/residual waste, has no accepted organic route, or is outside the scope of a liner-only/positive-list rule.",
        "Unclear or not stated": "The application is mentioned, but the source does not provide enough application-specific evidence to classify it as accepted, conditionally accepted, or rejected.",
    }
    for i, cat in enumerate(["Accepted", "Accepted with conditions", "Rejected", "Unclear or not stated"], 1):
        rows.append({"section":"Application-level indicator", "indicator":"Application Decision", "question_answered":"What is the four-category application-level compatibility decision before scenario conversion?", "category":cat, "definition":app_decision_definitions[cat], "included_examples":CODEBOOK_INCLUDED_EXAMPLES.get(cat, "Application-level mapping example."), "excluded_examples":CODEBOOK_EXCLUDED_EXAMPLES.get(cat, "Other application decision classes."), "include_regex_examples":"Not regex-coded directly; derived from Acceptance, Acceptance Condition, Certification/Approval Basis, and Application Group.", "exclude_regex_examples":"Not regex-coded directly; excluded by the derived decision mapping.", "regex_or_rule_examples":"Rule mapping from Acceptance + Acceptance Condition + Certification/Approval Basis + Application Group; see decision_basis in Application_Decisions.", "coding_basis":"derived_mapping", "priority_or_mapping_note":f"Priority/mapping class {i}; application-specific rejection and operational restrictions are applied before ordinary acceptance."})
    cert_sufficiency_definitions = {
        "Certification accepted": "The source-level Certification and Approval Basis is a certification, standard, or official approval category, and the source × application-group record is coded Accepted.",
        "Certification conditionally accepted": "The source-level Certification and Approval Basis is a certification, standard, or official approval category, and the source × application-group record is coded Accepted with conditions. Certification is relevant, but additional operational gatekeeping remains.",
        "Certification rejected": "The source-level Certification and Approval Basis is a certification, standard, or official approval category, but the source × application-group record is coded Rejected.",
        "Certification unclear or not stated": "The source-level Certification and Approval Basis is a certification, standard, or official approval category, but the source × application-group record is coded Unclear or not stated.",
    }
    for i, cat in enumerate(CERTIFICATION_SUFFICIENCY_ORDER, 1):
        rows.append({
            "section": "Derived certification indicator",
            "indicator": "Certification Sufficiency",
            "question_answered": "Among application records from sources that mention a certification, standard, or official approval basis, does certification correspond to acceptance, conditional acceptance, rejection, or uncertainty?",
            "category": cat,
            "definition": cert_sufficiency_definitions[cat],
            "included_examples": CODEBOOK_INCLUDED_EXAMPLES.get(cat, cert_sufficiency_definitions[cat]),
            "excluded_examples": CODEBOOK_EXCLUDED_EXAMPLES.get(cat, "Records whose source-level Certification and Approval Basis is generic compostable/biodegradable wording only or no named standard/approval stated are excluded from this denominator."),
            "include_regex_examples": "No direct include regex. First filter to certification-relevant Certification and Approval Basis categories: EN 13432 / OK compost / Seedling; Government/programme approval; BPI / ASTM / CMA; OK compost HOME / NF T 51-800 / AS 5810; Country-specific standards/certifications (BNQ / AS 4736 / DINplus). Then map Application Decision.",
            "exclude_regex_examples": "Generic compostable/biodegradable wording only and No named standard/approval stated are excluded before Certification Sufficiency is derived.",
            "regex_or_rule_examples": "Derived mapping implemented in certification_sufficiency_status(): Accepted => Certification accepted; Accepted with conditions => Certification conditionally accepted; Rejected => Certification rejected; Unclear or not stated => Certification unclear or not stated. Mapping is applied only after filtering to certification-relevant Certification and Approval Basis categories.",
            "coding_basis": "derived_mapping_after_denominator_filter",
            "priority_or_mapping_note": f"Mapping class {i}; derived after source-level Certification and Approval Basis and source × application-group Application Decision coding. Denominator is certification-relevant source × application-group records only.",
        })

    scenario_map = {
        "Lower-bound compatible share":"Accepted = compatible; Accepted with conditions = incompatible; Rejected = incompatible; Unclear/not stated excluded from denominator.",
        "Central compatible share":"Accepted and Accepted with conditions = compatible; Rejected = incompatible; Unclear/not stated excluded from denominator.",
        "Upper-bound compatible share":"Accepted, Accepted with conditions, and Unclear/not stated = compatible; Rejected = incompatible.",
    }
    for cat, rule in scenario_map.items():
        rows.append({"section":"Application-level scenario", "indicator":"Application-level Compatibility Scenario", "question_answered":"How are four-category application decisions converted into compatibility shares?", "category":cat, "definition":rule, "included_examples":CODEBOOK_INCLUDED_EXAMPLES.get(cat, "Scenario-compatible decisions described in the mapping rule."), "excluded_examples":CODEBOOK_EXCLUDED_EXAMPLES.get(cat, "Scenario-incompatible or denominator-excluded decisions described in the mapping rule."), "include_regex_examples":rule, "exclude_regex_examples":"Scenario-incompatible decisions are stated in the mapping rule.", "regex_or_rule_examples":rule, "coding_basis":"derived_mapping", "priority_or_mapping_note":"Scenario mapping; not directly regex-coded."})
    src_group_map = {
        "National/supranational sources":"National rule or policy + Supranational framework.",
        "Sources below national level":"Municipal sorting rule + waste operator/facility rule + state/provincial/regional/canton rule + intermunicipal/regional waste authority.",
    }
    for cat, rule in src_group_map.items():
        rows.append({"section":"Source-level sensitivity", "indicator":"Source-level Sensitivity Group", "question_answered":"Do source-level results differ by source level?", "category":cat, "definition":rule, "included_examples":CODEBOOK_INCLUDED_EXAMPLES.get(cat, rule), "excluded_examples":CODEBOOK_EXCLUDED_EXAMPLES.get(cat, "Other source-authority groups."), "include_regex_examples":rule, "exclude_regex_examples":"Other source-authority groups are excluded.", "regex_or_rule_examples":rule, "coding_basis":"derived_from_source_authority", "priority_or_mapping_note":"Derived after Source Authority coding; not independently text-coded."})
    return rows


DISPLAY_LABELS: Dict[str, Dict[str, str]] = {
    "Stated Organic-waste Treatment Route": {
        "AD + composting": "AD/biogas + composting",
        "AD / biogas": "AD/biogas only",
        "Other valorisation": "Other organic recovery",
        "Treatment route not stated": "Treatment route not stated",
    },
    "Acceptance": {
        "Broad acceptance": "Accepted broadly",
        "Liners only": "Collection liners only",
        "Listed items only": "Positive-list / listed items only",
        "Controlled acceptance": "Dedicated/controlled route only",
        "No accepted route": "No clear organics route",
        "No explicit rule": "No explicit compostable-packaging rule",
        "Locally variable": "Local decision required",
    },
    "Acceptance Condition": {
        "Certification required": "Named certification/standard required",
        "Generic compostability/biodegradability requirement": "Generic compostable/biodegradable requirement",
        "Context-specific only": "Controlled/dedicated context required",
        "No condition / not applicable": "No additional condition stated",
        "Not applicable - not accepted/no explicit route": "Not applicable - not accepted/no explicit rule",
    },
    "Certification or Approval Basis": {
        "Generic compostable/biodegradable claim only": "Generic compostable/biodegradable wording only",
        "No standard stated": "No named standard/approval stated",
        "Government/programme-approved": "Government/programme approval",
        "BNQ / AS 4736 / DINplus / national standards": "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
        "National standards (BNQ / AS 4736 / DINplus etc.)": "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
        "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)": "Country-specific standards/certifications (BNQ / AS 4736 / DINplus)",
    },
    "Rejection Rationale": {
        "Explicit no-packaging / positive-list restriction": "Explicit no-packaging / positive-list restriction",
        "Compost/digestate quality concern": "Compost/digestate quality or contamination concern",
        "No compatible treatment route": "No compatible organic-treatment route",
        "Reason not specified": "Reason not explicitly stated",
        "Not applicable": "Not applicable - not a rejection/restriction case",
    },
}


def map_indicator_value(indicator: str, value: Any) -> Any:
    if pd.isna(value):
        return value
    return DISPLAY_LABELS.get(indicator, {}).get(str(value), value)


def apply_display_labels(tables: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    out = {k: v.copy() for k, v in tables.items()}
    ref_cols = [
        "Stated Organic-waste Treatment Route", "Acceptance",
        "Certification or Approval Basis", "Rejection Rationale",
    ]
    if "reference_level_coding_display" in out:
        for col in ref_cols:
            if col in out["reference_level_coding_display"].columns:
                out["reference_level_coding_display"][col] = out["reference_level_coding_display"][col].map(lambda x, c=col: map_indicator_value(c, x))
    for key in ["reference_level_category_shares_display", "rejection_rationale_reporting_shares_display", "source_level_sensitivity_display"]:
        if key in out and {"indicator", "category"}.issubset(out[key].columns):
            out[key]["category"] = out[key].apply(lambda r: map_indicator_value(r["indicator"], r["category"]), axis=1)
            if "denominator_basis" in out[key].columns:
                out[key]["denominator_basis"] = out[key]["denominator_basis"].astype(str).replace({
                    "References with accepted or conditionally accepted route; excludes rejection/no-route/no-explicit-rule cases": "References with accepted or conditionally accepted route; excludes rejected/no-route/no-explicit-rule cases",
                    "References with specified rejection/restriction rationale; excludes Not applicable and Reason not specified": "References with specified rejection/restriction rationale; excludes not-applicable and reason-not-explicitly-stated cases",
                })
    if "codebook_display" in out and {"indicator", "category"}.issubset(out["codebook_display"].columns):
        out["codebook_display"]["category"] = out["codebook_display"].apply(lambda r: map_indicator_value(r["indicator"], r["category"]), axis=1)
    if "reference_rule_trace_audit" in out and {"indicator", "assigned_category"}.issubset(out["reference_rule_trace_audit"].columns):
        out["reference_rule_trace_audit"]["assigned_category"] = out["reference_rule_trace_audit"].apply(lambda r: map_indicator_value(r["indicator"], r["assigned_category"]), axis=1)
    if "application_rule_trace_audit" in out and "reference_acceptance" in out["application_rule_trace_audit"].columns:
        out["application_rule_trace_audit"]["reference_acceptance"] = out["application_rule_trace_audit"]["reference_acceptance"].map(lambda x: map_indicator_value("Acceptance", x))

    cert_long_key = "certification_sufficiency_long_display"
    if cert_long_key in out and not out[cert_long_key].empty:
        for col in ["Acceptance", "Acceptance Condition", "Certification or Approval Basis", "Rejection Rationale"]:
            if col in out[cert_long_key].columns:
                out[cert_long_key][col] = out[cert_long_key][col].map(lambda x, c=col: map_indicator_value(c, x))
    cert_by_key = "certification_sufficiency_by_certification_basis_display"
    if cert_by_key in out and "Certification or Approval Basis" in out[cert_by_key].columns:
        out[cert_by_key]["Certification or Approval Basis"] = out[cert_by_key]["Certification or Approval Basis"].map(lambda x: map_indicator_value("Certification or Approval Basis", x))

    # Align output-table terminology with the manuscript structure.
    # Do not remove internal/intermediate coding indicators here; workbook-facing
    # sheet filtering belongs in the separate Excel-generation script.
    indicator_rename = {
        "Certification or Approval Basis": "Certification and Approval Basis",
        "Certification Sufficiency / Accountability Gap": "Certification Sufficiency",
        "Binary Compatibility Scenario": "Application-level Compatibility Scenario",
        "Source-level sensitivity group": "Source-level Sensitivity Group",
    }
    for key, df in list(out.items()):
        if df is None or df.empty:
            continue
        if "indicator" in df.columns:
            df["indicator"] = df["indicator"].replace(indicator_rename)
        df.rename(columns={
            "Certification or Approval Basis": "Certification and Approval Basis",
            "Source-level sensitivity group": "Source-level Sensitivity Group",
        }, inplace=True)
        if "denominator_basis" in df.columns:
            df["denominator_basis"] = df["denominator_basis"].astype(str).str.replace("reference × detected application group", "source × application group", regex=False)
            df["denominator_basis"] = df["denominator_basis"].astype(str).str.replace("Reference × detected application group", "Source × application group", regex=False)
            df["denominator_basis"] = df["denominator_basis"].astype(str).str.replace("references", "sources", regex=False)
            df["denominator_basis"] = df["denominator_basis"].astype(str).str.replace("References", "Sources", regex=False)
        for col in [c for c in df.columns if is_object_dtype(df[c].dtype) or is_string_dtype(df[c].dtype)]:
            df[col] = (df[col].astype(str)
                .str.replace("Certification or Approval Basis", "Certification and Approval Basis", regex=False)
                .str.replace("Certification/Approval Basis", "Certification and Approval Basis", regex=False)
                .str.replace("Acceptance Condition", "conditional-acceptance cue", regex=False)
                .str.replace("Application Group", "Application Group", regex=False)
                .str.replace("application group", "application group", regex=False)
                .str.replace("Application Groups", "Application Groups", regex=False)
                .str.replace("application groups", "application groups", regex=False)
                .str.replace("Certification Sufficiency / Accountability Gap", "Certification Sufficiency", regex=False)
                .str.replace("certification-sufficiency", "certification-sufficiency", regex=False)
                .str.replace("certification sufficiency", "certification sufficiency", regex=False)
                .str.replace("Reference × detected application group", "Source × application group", regex=False)
                .str.replace("reference × detected application group", "source × application group", regex=False)
                .str.replace("Reference-level", "Source-level", regex=False)
                .str.replace("reference-level", "source-level", regex=False)
                .str.replace("retained references", "included sources", regex=False)
                .str.replace("retained reference", "included source", regex=False))
    return out


def build_outputs(rows: List[Dict[str, Any]]) -> Dict[str, pd.DataFrame]:
    reference_rows: List[Dict[str, Any]] = []
    app_rows: List[Dict[str, Any]] = []
    ref_audit: List[Dict[str, Any]] = []
    app_audit: List[Dict[str, Any]] = []

    for src in rows:
        cid = src.get("citation_id", "")
        label = src.get("source_label", "")
        title = src.get("source_title", "")
        url = src.get("source_url", "")
        country, region = source_label_parts(label)
        full_text = src.get("text", "") or ""
        win = evidence_windows(full_text)
        meta = meta_text(src)
        combined = "\n".join([meta, win])

        source_hit = classify_source_authority(src)
        # Acceptance is the most nuance-sensitive indicator; scan full source text
        # plus metadata to avoid missing body text on navigation-heavy municipal pages.
        acceptance_text = "\n".join([meta, full_text])
        acceptance_hit = classify_acceptance(acceptance_text)
        treatment_hit = classify_treatment_route(combined)
        treatment_note = treatment_route_coding_note(treatment_hit)
        cert_text = "\n".join([meta, full_text])
        cert_hit = classify_certification_basis(cert_text)
        cond_hit = classify_acceptance_condition(combined, acceptance_hit.category, cert_hit.category)
        rationale_hit = classify_rejection_rationale(acceptance_text, acceptance_hit.category)
        group = source_level_group(source_hit.category)
        primary = primary_for_sensitivity(source_hit.category)
        app_hits = detect_application_types(combined)

        ref = {
            "citation_id": cid,
            "country_or_region": country,
            "subnational_or_context": region,
            "source_label": label,
            "source_title": title,
            "source_url": url,
            "Source Authority": source_hit.category,
            "Stated Organic-waste Treatment Route": treatment_hit.category,
            "Treatment Route Coding Note": treatment_note,
            "Acceptance": acceptance_hit.category,
            "Certification or Approval Basis": cert_hit.category,
            "Rejection Rationale": rationale_hit.category,
            "Application Groups Detected": "; ".join([h.category for h in app_hits]),
            "Source-level sensitivity group": group,
        }
        reference_rows.append(ref)

        for indicator, hit in [
            ("Source Authority", source_hit),
            ("Stated Organic-waste Treatment Route", treatment_hit),
            ("Acceptance", acceptance_hit),
            ("Certification or Approval Basis", cert_hit),
            ("Rejection Rationale", rationale_hit),
        ]:
            ref_audit.append({
                "citation_id": cid,
                "source_label": label,
                "indicator": indicator,
                "assigned_category": hit.category,
                "matched_regex_or_rule": hit.matched_rule,
                "matched_text": hit.matched_text,
                "evidence_excerpt": hit.excerpt,
                "coding_basis": hit.coding_basis,
            })

        for app_hit in app_hits:
            if app_hit.category == "Organic-waste system context only":
                continue
            decision, basis = decide_application(app_hit.category, acceptance_hit.category, combined, cond_hit.category, cert_hit.category)
            app_rows.append({
                "citation_id": cid,
                "country_or_region": country,
                "source_label": label,
                "source_title": title,
                "source_url": url,
                "application_type": app_hit.category,
                "application_decision": decision,
                "decision_basis": basis,
            })
            app_audit.append({
                "citation_id": cid,
                "source_label": label,
                "application_type": app_hit.category,
                "application_decision": decision,
                "decision_basis": basis,
                "matched_regex_or_rule": app_hit.matched_rule,
                "matched_text": app_hit.matched_text,
                "evidence_excerpt": app_hit.excerpt,
                "reference_acceptance": acceptance_hit.category,
                "reference_acceptance_matched_rule": acceptance_hit.matched_rule,
                "reference_acceptance_matched_text": acceptance_hit.matched_text,
            })

    ref_df = pd.DataFrame(reference_rows)
    app_df = pd.DataFrame(app_rows)
    ref_audit_df = pd.DataFrame(ref_audit)
    app_audit_df = pd.DataFrame(app_audit)

    shares = pd.concat([
        share_table(ref_df, "Source Authority", "Source Authority", "What kinds of sources support the source corpus?", "All included sources"),
        share_table(ref_df, "Stated Organic-waste Treatment Route", "Stated Organic-waste Treatment Route", "What organic-waste collection or treatment route is stated?", "All included sources"),
        share_table(ref_df, "Acceptance", "Acceptance", "How are compostable packaging applications treated in organic-waste systems?", "All included sources"),
        share_table(ref_df, "Certification or Approval Basis", "Certification or Approval Basis", "What standard, label, programme, or approval basis is mentioned?", "All included sources"),
    ], ignore_index=True)

    # Main reporting for rejection rationales uses only specified rationales.
    # "Reason not specified" and "Not applicable" are retained in Reference_Coding
    # and validation/audit outputs, but are excluded from the main reason-share denominator.
    specified_rationale_mask = ~ref_df["Rejection Rationale"].isin(["Not applicable", "Reason not specified"])
    rejection_reporting = share_table(
        ref_df,
        "Rejection Rationale",
        "Rejection Rationale",
        "Among references with a specified rejection/restriction rationale, what reasons are stated?",
        "References with specified rejection/restriction rationale; excludes Not applicable and Reason not specified",
        mask=specified_rationale_mask,
    )
    shares = pd.concat([shares, rejection_reporting], ignore_index=True)
    treatment_audit = treatment_route_not_stated_audit(ref_df)
    if not treatment_audit.empty:
        treatment_audit_summary = (
            treatment_audit.groupby("Treatment Route Coding Note")
            .size()
            .reset_index(name="n")
            .sort_values("n", ascending=False)
        )
        treatment_audit_summary["denominator"] = int(treatment_audit_summary["n"].sum())
        treatment_audit_summary["share"] = treatment_audit_summary["n"] / treatment_audit_summary["denominator"].replace({0: pd.NA})
        treatment_audit_summary["share_percent"] = 100 * treatment_audit_summary["share"]
        treatment_audit_summary["denominator_basis"] = "References coded as Treatment route not stated."
    else:
        treatment_audit_summary = pd.DataFrame(columns=["Treatment Route Coding Note", "n", "denominator", "share", "share_percent", "denominator_basis"])

    if app_df.empty:
        app_shares = pd.DataFrame(columns=["application_type", "application_decision", "n", "denominator", "share", "share_percent", "denominator_basis"])
        app_compat = pd.DataFrame()
    else:
        app_counts = app_df.groupby(["application_type", "application_decision"]).size().reset_index(name="n")
        den = app_df.groupby("application_type").size().rename("denominator").reset_index()
        app_shares = app_counts.merge(den, on="application_type", how="left")
        app_shares["share"] = app_shares["n"] / app_shares["denominator"]
        app_shares["share_percent"] = 100 * app_shares["share"]
        app_shares["denominator_basis"] = "Application-decision records for the same application group (reference × detected application group)."
        comp_rows = []
        for app_type in APPLICATION_ORDER:
            sub = app_df[app_df["application_type"] == app_type]
            counts = sub["application_decision"].value_counts().to_dict()
            comp_rows.append(binary_scenario_row(app_type, counts))
        app_compat = pd.DataFrame(comp_rows)
        app_compat["has_zero_records"] = app_compat["evidence_records"].eq(0).map({True: "Yes", False: "No"})
        app_compat["denominator_basis"] = "Application-decision records for the same application group (source × detected application group). Lower/central scenarios exclude unclear/not-stated cases from the denominator; upper scenario includes them."

    cert_sufficiency_tables = build_certification_sufficiency_tables(ref_df, app_df)

    sens_rows = []
    ref_indicators = [
        ("Stated Organic-waste Treatment Route", "Stated Organic-waste Treatment Route"),
        ("Acceptance", "Acceptance"),
        ("Certification or Approval Basis", "Certification or Approval Basis"),
        ("Rejection Rationale", "Rejection Rationale"),
    ]
    groups = ["All retained sources", "National/supranational sources", "Sources below national level"]
    primary_df = ref_df.copy()
    for indicator, col in ref_indicators:
        for grp in groups:
            if grp == "All retained sources":
                gdf = primary_df
            else:
                gdf = primary_df[primary_df["Source-level sensitivity group"] == grp]
            if indicator == "Rejection Rationale":
                gdf = gdf[~gdf[col].isin(["Not applicable", "Reason not specified"])]
            denom = len(gdf)
            for cat, n in gdf[col].value_counts(dropna=False).items():
                sens_rows.append({
                    "indicator": indicator,
                    "source_level_group": grp,
                    "category": cat,
                    "n": int(n),
                    "denominator": int(denom),
                    "share": n / denom if denom else 0,
                    "share_percent": 100 * n / denom if denom else 0,
                })
    sensitivity = pd.DataFrame(sens_rows)

    validation = []
    validation.append({"check": "included_sources", "value": len(ref_df), "status": "Info", "note": "Number of JSONL sources coded."})
    validation.append({"check": "application_decision_rows", "value": len(app_df), "status": "Info", "note": "Number of source × application-group rows."})
    zero_apps = [] if app_compat.empty else app_compat.loc[app_compat["evidence_records"] == 0, "application_type"].tolist()
    validation.append({"check": "application_categories_with_zero_records", "value": len(zero_apps), "status": "Pass" if len(zero_apps) == 0 else "Review", "note": "; ".join(zero_apps)})
    valid_sensitivity_groups = {"National/supranational sources", "Sources below national level"}
    n_allocated = int(ref_df["Source-level sensitivity group"].isin(valid_sensitivity_groups).sum()) if "Source-level sensitivity group" in ref_df.columns else 0
    validation.append({"check": "sources_allocated_to_sensitivity", "value": n_allocated, "status": "Pass" if n_allocated == len(ref_df) else "Review", "note": "All included sources should be allocated to a national/supranational source group or a sources-below-national-level group."})
    validation.append({"check": "rejection_rationale_reason_not_specified", "value": int((ref_df["Rejection Rationale"] == "Reason not specified").sum()), "status": "Info", "note": "Audit count only; excluded from the main specified-rationale denominator."})
    validation.append({"check": "rejection_rationale_not_applicable", "value": int((ref_df["Rejection Rationale"] == "Not applicable").sum()), "status": "Info", "note": "Sources not coded as rejection/restriction cases; excluded from rejection-rationale reporting."})
    validation.append({"check": "certification_sufficiency_rows", "value": len(cert_sufficiency_tables["certification_sufficiency_long_display"]), "status": "Info", "note": "Certification-relevant source × application-group rows used for certification-sufficiency analysis; generic wording and no-standard records are excluded."})
    validation_df = pd.DataFrame(validation)

    tables = {
        "reference_level_coding_display": ref_df,
        "reference_level_category_shares_display": shares,
        "treatment_route_not_stated_audit_display": treatment_audit,
        "treatment_route_not_stated_audit_summary_display": treatment_audit_summary,
        "rejection_rationale_reporting_shares_display": rejection_reporting,
        "application_decision_long_display": app_df,
        "application_decision_shares_display": app_shares,
        "application_compatibility_scenarios_display": app_compat,
        **cert_sufficiency_tables,
        "source_level_sensitivity_display": sensitivity,
        "codebook_display": pd.DataFrame(codebook_rows()),
        "validation_checks_display": validation_df,
        "reference_rule_trace_audit": ref_audit_df,
        "application_rule_trace_audit": app_audit_df,
        "application_category_coverage_audit": app_compat[["application_type", "evidence_records", "has_zero_records"]].copy() if not app_compat.empty else pd.DataFrame(),
    }
    return apply_display_labels(tables)


OUTPUT_FILE_NAME_MAP = {
    "reference_level_coding_display": "source_level_coding_display.csv",
    "reference_level_category_shares_display": "source_level_category_shares_display.csv",
    "reference_rule_trace_audit": "source_rule_trace_audit.csv",
}


def write_outputs(tables: Dict[str, pd.DataFrame], outdir: Path, zip_outputs: bool = False) -> None:
    if outdir.exists():
        shutil.rmtree(outdir)
    display_dir = outdir / "display"
    audit_dir = outdir / "audit"
    display_dir.mkdir(parents=True, exist_ok=True)
    audit_dir.mkdir(parents=True, exist_ok=True)

    # Reader-facing display CSVs: these directly support the Excel workbook,
    # manuscript figures, or Supplementary Data. Long trace/QA files are kept
    # under audit/ to keep display/ clean.
    display_keys = [
        "reference_level_coding_display",
        "reference_level_category_shares_display",
        "treatment_route_not_stated_audit_summary_display",
        "rejection_rationale_reporting_shares_display",
        "application_decision_long_display",
        "application_decision_shares_display",
        "application_compatibility_scenarios_display",
        "certification_sufficiency_summary_display",
        "certification_sufficiency_by_application_display",
        "certification_sufficiency_bubble_matrix_display",
        "certification_sufficiency_by_certification_basis_display",
        "certification_sufficiency_by_source_authority_display",
        "certification_sufficiency_long_display",
        "source_level_sensitivity_display",
        "codebook_display",
    ]

    # Audit CSVs: detailed traceability, QA, or diagnostic files. These are
    # useful for checking the coding workflow but are not part of the main
    # reader-facing result set.
    audit_keys = [
        "treatment_route_not_stated_audit_display",
        "validation_checks_display",
        "reference_rule_trace_audit",
        "application_rule_trace_audit",
        "application_category_coverage_audit",
    ]
    for key in display_keys:
        filename = OUTPUT_FILE_NAME_MAP.get(key, f"{key}.csv")
        tables[key].to_csv(display_dir / filename, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
    for key in audit_keys:
        filename = OUTPUT_FILE_NAME_MAP.get(key, f"{key}.csv")
        tables[key].to_csv(audit_dir / filename, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)

    report = outdir / "audit_report.md"
    with report.open("w", encoding="utf-8") as f:
        f.write("# Rule-core coding audit report\n\n")
        f.write(f"Generated: {dt.datetime.now().isoformat(timespec='seconds')}\n\n")
        for _, row in tables["validation_checks_display"].iterrows():
            f.write(f"- **{row['check']}**: {row['value']} ({row['status']}) — {row['note']}\n")

    if zip_outputs:
        for suffix, folder in [("display", display_dir), ("audit", audit_dir), ("full", outdir)]:
            zip_path = outdir.parent / f"{outdir.name}_{suffix}.zip"
            if zip_path.exists():
                zip_path.unlink()
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                if suffix == "full":
                    for p in outdir.rglob("*"):
                        if p.is_file():
                            zf.write(p, p.relative_to(outdir.parent))
                else:
                    for p in folder.rglob("*"):
                        if p.is_file():
                            zf.write(p, p.relative_to(outdir.parent))


def main() -> None:
    parser = argparse.ArgumentParser(description="Calculate compostable-packaging operational-rule results from the archived evidence corpus.")
    parser.add_argument("--input", default=str(INPUT_ZIP_DEFAULT), help="Path to web_sources.zip")
    parser.add_argument("--outdir", default=str(OUTDIR_DEFAULT), help="Output directory")
    parser.add_argument("--zip", action="store_true", help="Optionally create display/audit/full zip archives")
    args = parser.parse_args()

    input_zip = Path(args.input).expanduser().resolve()
    outdir = Path(args.outdir).expanduser().resolve()
    rows = load_corpus(input_zip)
    tables = build_outputs(rows)
    write_outputs(tables, outdir, zip_outputs=args.zip)

    print("Result calculation complete.")
    print(f"Sources coded: {len(tables['reference_level_coding_display'])}")
    print(f"Application decision rows: {len(tables['application_decision_long_display'])}")
    if not tables['application_compatibility_scenarios_display'].empty:
        zero = tables['application_compatibility_scenarios_display'].query("evidence_records == 0")['application_type'].tolist()
        print(f"Application categories with zero records: {len(zero)}")
        if zero:
            print("Zero categories:", "; ".join(zero))
    print(f"Outputs written to: {outdir}")


if __name__ == "__main__":
    main()
