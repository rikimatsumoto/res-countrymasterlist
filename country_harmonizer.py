"""
Country Harmonizer — Master Country List Manager
==================================================
A tool for harmonizing inconsistent country names across datasets (IMF, World Bank,
UN, ILO, etc.) using a single master Excel file as the source of truth.

Core capabilities:
  1. LOOKUP    — Find a country by *any* key: name variant, ISO2, ISO3, UN code, IFS code
  2. FUZZY     — Handle typos and close variants via fuzzy string matching
  3. HARMONIZE — Auto-standardize a DataFrame column of messy country names in one call
  4. UPDATE    — Pull fresh classifications from the World Bank and IMF APIs
  5. VERSION   — Track every change with timestamped backups + optional Git integration

Usage:
  # As a module in your workflow
  from country_harmonizer import CountryHarmonizer
  ch = CountryHarmonizer("MasterCountryList.xlsx")
  clean_df = ch.harmonize(df, country_col="Country Name")

  # As a CLI to update the master file from the World Bank API
  python country_harmonizer.py update --file MasterCountryList.xlsx
"""

import json
import shutil
import hashlib
import logging
import argparse
import subprocess
from pathlib import Path
from datetime import datetime
from typing import Optional

import pandas as pd
import urllib.request
import urllib.error


# ===========================================================================
# FUZZY MATCHING — Pure Python (replaces thefuzz / python-Levenshtein)
# ===========================================================================
# We only need two things from thefuzz:
#   1. A string similarity scorer (token_sort_ratio)
#   2. A "find best match from a list" function (extractOne)
#
# Both are built on Levenshtein edit distance, which is ~15 lines of Python.

def _levenshtein(s1: str, s2: str) -> int:
    """
    Compute the Levenshtein edit distance between two strings.

    This is the minimum number of single-character edits (insertions,
    deletions, substitutions) needed to transform s1 into s2.

    Uses the standard dynamic-programming approach with O(min(m,n)) space
    by only keeping two rows of the matrix at a time.
    """
    # Ensure s1 is the shorter string (saves memory)
    if len(s1) > len(s2):
        s1, s2 = s2, s1

    # prev_row[j] = edit distance between s1[:0] and s2[:j] = j
    prev_row = list(range(len(s2) + 1))

    for i, c1 in enumerate(s1):
        # curr_row[0] = edit distance between s1[:i+1] and s2[:0] = i+1
        curr_row = [i + 1]
        for j, c2 in enumerate(s2):
            # Three options: insert, delete, or substitute
            insert  = prev_row[j + 1] + 1
            delete  = curr_row[j] + 1
            replace = prev_row[j] + (0 if c1 == c2 else 1)
            curr_row.append(min(insert, delete, replace))
        prev_row = curr_row

    return prev_row[-1]


def _token_sort_ratio(s1: str, s2: str) -> int:
    """
    Score similarity between two strings (0–100), invariant to word order.

    This replicates thefuzz's token_sort_ratio:
      1. Lowercase both strings
      2. Split into tokens, sort alphabetically, rejoin
      3. Compute Levenshtein ratio on the sorted versions

    "Korea, Rep." vs "Rep. Korea" → same tokens → high score.
    """
    # Normalize: lowercase, strip punctuation, sort tokens
    def _normalize(s):
        # Replace punctuation with spaces (so "Timor-Leste" → "timor leste")
        cleaned = "".join(c if c.isalnum() or c.isspace() else " " for c in s)
        tokens = sorted(cleaned.lower().split())
        return " ".join(tokens)

    norm1 = _normalize(s1)
    norm2 = _normalize(s2)

    # Levenshtein ratio = 1 - (distance / max_length), scaled to 0–100
    if not norm1 and not norm2:
        return 100
    max_len = max(len(norm1), len(norm2))
    if max_len == 0:
        return 100

    dist = _levenshtein(norm1, norm2)
    return int(round((1 - dist / max_len) * 100))


def _find_best_match(query: str, candidates: list) -> tuple:
    """
    Find the best fuzzy match for `query` from a list of candidates.

    Replaces thefuzz's process.extractOne(). Scans all candidates and
    returns the (best_match_string, score) tuple.

    Args:
        query: The string to match
        candidates: List of strings to match against

    Returns:
        (best_match, score) — e.g., ("Russia", 92)
    """
    best_match = ""
    best_score = 0

    for candidate in candidates:
        score = _token_sort_ratio(query, candidate)
        if score > best_score:
            best_score = score
            best_match = candidate

    return (best_match, best_score)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Minimum score (0–100) a fuzzy match must reach to be accepted automatically.
# Matches below this are flagged for manual review.
FUZZY_THRESHOLD = 85

# How many fuzzy candidates to consider when matching
FUZZY_CANDIDATES = 5

# World Bank API base URL (v2, JSON format)
WB_API_BASE = "https://api.worldbank.org/v2"

# Columns in the master file that come from the World Bank and can be auto-updated
WB_UPDATEABLE_COLS = [
    "wb_country", "wb_isocode", "wb_region", "wb_incomegroup",
    "wb_lendingcategory",
]

# IMF DataMapper API base URL
IMF_API_BASE = "https://www.imf.org/external/datamapper/api/v1"

# IMF group IDs → master file 'income' column values
# The DataMapper uses short codes for its analytical groups. We map them
# to the human-readable labels used in your master file.
IMF_INCOME_GROUPS = {
    "ADVEC":  "Advanced Economies",
    "OEMDC":  "Emerging Market Economies",     # "Emerging Market and Developing Economies"
    "LIDC":   "Low Income Developing Countries",
}

# Columns in the master file that come from the IMF and can be auto-updated
IMF_UPDATEABLE_COLS = ["region", "income"]

# Directory name for backups (created next to the master file)
BACKUP_DIR = "backups"

# Changelog filename (created next to the master file)
CHANGELOG_FILE = "country_harmonizer_changelog.json"

# Set up module-level logger
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("country_harmonizer")


# ===========================================================================
# STEP 1 — The Lookup Index
# ===========================================================================
# We build a flat dictionary that maps EVERY known name/code variant → isocode.
# This is what makes "lookup by any key" possible in O(1).

def _build_lookup_index(df: pd.DataFrame) -> dict:
    """
    Build a case-insensitive reverse lookup from every name/code variant
    to the canonical ISO3 code.

    For each country row, we index:
      - country (IMF name)
      - wb_country (World Bank name)
      - country_french (UN French name)
      - isocode (ISO3 — e.g., 'USA')
      - isocode2 (ISO2 — e.g., 'US')
      - un_code (numeric UN code — e.g., 840)
      - ifscode (IFS numeric code)

    Returns:
        dict: {lowercase_variant: isocode, ...}
    """
    index = {}

    # Columns that contain name/code variants we want to index
    variant_cols = [
        "country", "wb_country", "country_french",
        "isocode", "isocode2", "un_code", "ifscode",
    ]

    # Columns with numeric codes that should also be indexed as integers
    numeric_cols = {"un_code", "ifscode"}

    for _, row in df.iterrows():
        iso3 = str(row["isocode"]).strip().upper()

        for col in variant_cols:
            val = row.get(col)
            if pd.notna(val):
                # Normalize: lowercase string, strip whitespace
                key = str(val).strip().lower()
                if key:
                    index[key] = iso3

                # For numeric columns, also index the integer form
                # (e.g., un_code 840.0 → index both "840.0" and "840")
                if col in numeric_cols:
                    try:
                        int_key = str(int(float(val)))
                        index[int_key] = iso3
                    except (ValueError, OverflowError):
                        pass

    log.info(f"Lookup index built: {len(index)} entries → {df.shape[0]} countries")
    return index


# ===========================================================================
# STEP 2 — Version Control Helpers
# ===========================================================================

def _is_git_repo(path: Path) -> bool:
    """Check if the given path is inside a Git repository."""
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--is-inside-work-tree"],
            cwd=str(path), capture_output=True, text=True, timeout=5,
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def _git_commit(path: Path, message: str):
    """Stage the master file and commit with the given message."""
    try:
        subprocess.run(["git", "add", str(path)], cwd=str(path.parent),
                        capture_output=True, check=True, timeout=10)
        subprocess.run(["git", "commit", "-m", message], cwd=str(path.parent),
                        capture_output=True, check=True, timeout=10)
        log.info(f"Git commit: {message}")
    except subprocess.CalledProcessError as e:
        log.warning(f"Git commit failed (maybe no changes?): {e.stderr}")


def _create_backup(filepath: Path) -> Path:
    """
    Create a timestamped backup of the master file.
    Returns the path to the backup.
    """
    backup_dir = filepath.parent / BACKUP_DIR
    backup_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = filepath.stem
    backup_path = backup_dir / f"{stem}_{timestamp}{filepath.suffix}"
    shutil.copy2(filepath, backup_path)

    log.info(f"Backup created: {backup_path.name}")
    return backup_path


def _append_changelog(filepath: Path, entry: dict):
    """Append a structured entry to the changelog JSON file."""
    changelog_path = filepath.parent / CHANGELOG_FILE
    history = []

    if changelog_path.exists():
        with open(changelog_path, "r") as f:
            try:
                history = json.load(f)
            except json.JSONDecodeError:
                history = []

    history.append(entry)

    with open(changelog_path, "w") as f:
        json.dump(history, f, indent=2, default=str)

    log.info(f"Changelog updated: {entry['action']}")


# ===========================================================================
# STEP 3 — World Bank API Fetcher
# ===========================================================================

def fetch_wb_countries() -> pd.DataFrame:
    """
    Fetch the full country list from the World Bank API.

    The WB API paginates, so we request a large page size (300) to get
    everything in one call. Returns a DataFrame with columns aligned
    to the master file's wb_* columns.

    Returns:
        pd.DataFrame with columns:
            wb_isocode, wb_country, wb_region, wb_incomegroup, wb_lendingcategory
    """
    url = f"{WB_API_BASE}/country?format=json&per_page=300"
    log.info(f"Fetching World Bank country data from: {url}")

    # Using stdlib urllib instead of requests — one fewer dependency
    req = urllib.request.Request(url)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            if resp.status != 200:
                raise RuntimeError(f"World Bank API returned status {resp.status}")
            raw = resp.read().decode("utf-8")
    except urllib.error.URLError as e:
        raise ConnectionError(f"Failed to reach World Bank API: {e}") from e

    data = json.loads(raw)

    # WB API returns [metadata_dict, list_of_countries]
    if len(data) < 2:
        raise ValueError("Unexpected World Bank API response format")

    records = []
    for entry in data[1]:
        records.append({
            "wb_isocode":          entry.get("id", ""),
            "wb_country":          entry.get("name", ""),
            "wb_region":           entry.get("region", {}).get("value", ""),
            "wb_incomegroup":      entry.get("incomeLevel", {}).get("value", ""),
            "wb_lendingcategory":  entry.get("lendingType", {}).get("value", ""),
        })

    wb_df = pd.DataFrame(records)

    # Filter out aggregate regions (they have region = "Aggregates")
    wb_df = wb_df[wb_df["wb_region"] != "Aggregates"].copy()

    log.info(f"Fetched {len(wb_df)} countries from World Bank API")
    return wb_df


def _fetch_imf_json(endpoint: str) -> dict:
    """
    Fetch a JSON response from the IMF DataMapper API.

    Args:
        endpoint: Path after the base URL (e.g., "countries", "groups", "regions")

    Returns:
        Parsed JSON as a dict
    """
    url = f"{IMF_API_BASE}/{endpoint}"
    req = urllib.request.Request(url)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            if resp.status != 200:
                raise RuntimeError(f"IMF API returned status {resp.status}")
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as e:
        raise ConnectionError(f"Failed to reach IMF API ({url}): {e}") from e


def fetch_imf_classifications() -> dict:
    """
    Fetch country classifications from the IMF DataMapper API.

    The IMF API returns group-centric data:
        { "ADVEC": { "label": "Advanced Economies", "countries": ["USA", "GBR", ...] } }

    We invert this to country-centric:
        { "USA": { "region": "...", "income": "..." }, ... }

    Process:
      1. Fetch /groups  — analytical groups like "Advanced Economies" → maps to 'income'
      2. Fetch /regions — geographic regions like "Sub-Saharan Africa" → maps to 'region'
      3. Invert both from {group → [countries]} to {country → classification}

    Returns:
        dict: {iso3: {"region": "...", "income": "..."}, ...}
    """
    log.info("Fetching IMF classifications from DataMapper API...")

    country_data = {}  # {iso3: {"region": ..., "income": ...}}

    # --- Step 1: Fetch income groups ---
    # The /groups endpoint returns ALL analytical groups (G7, EU, ASEAN-5, etc.)
    # We only care about the three that map to the 'income' column.
    groups_raw = _fetch_imf_json("groups")

    for group_id, income_label in IMF_INCOME_GROUPS.items():
        group_info = groups_raw.get("groups", {}).get(group_id, {})
        countries = group_info.get("countries", [])
        for iso3 in countries:
            country_data.setdefault(iso3, {})
            country_data[iso3]["income"] = income_label

    log.info(f"  Income groups: mapped {len(country_data)} countries")

    # --- Step 2: Fetch regional groupings ---
    # The /regions endpoint structure: {region_id: {label: "...", countries: [...]}}
    regions_raw = _fetch_imf_json("regions")

    region_count = 0
    for region_id, region_info in regions_raw.get("regions", {}).items():
        region_label = region_info.get("label", region_id)
        countries = region_info.get("countries", [])
        for iso3 in countries:
            country_data.setdefault(iso3, {})
            country_data[iso3]["region"] = region_label
            region_count += 1

    log.info(f"  Regions: mapped {region_count} country-region pairs")
    log.info(f"  Total: {len(country_data)} countries with IMF classifications")

    return country_data


# ===========================================================================
# STEP 4 — The Main Harmonizer Class
# ===========================================================================

class CountryHarmonizer:
    """
    Central class for country name harmonization and master file management.

    Workflow:
        1. Load master Excel → build lookup index
        2. Use .lookup() / .harmonize() to match messy names
        3. Use .update_from_wb() to refresh World Bank classifications
        4. Use .save() to write changes with version control
    """

    def __init__(self, filepath: str, fuzzy_threshold: int = FUZZY_THRESHOLD):
        """
        Initialize the harmonizer by loading the master country list.

        Args:
            filepath: Path to the master Excel file (.xlsx)
            fuzzy_threshold: Minimum fuzzy match score (0-100) to auto-accept
        """
        self.filepath = Path(filepath).resolve()
        self.fuzzy_threshold = fuzzy_threshold

        # -------------------------------------------------------------------
        # Step 1a: Load the master data into a DataFrame
        # -------------------------------------------------------------------
        if not self.filepath.exists():
            raise FileNotFoundError(f"Master file not found: {self.filepath}")

        self.df = pd.read_excel(self.filepath, sheet_name="Main")
        log.info(f"Loaded master file: {self.df.shape[0]} countries, "
                 f"{self.df.shape[1]} columns")

        # -------------------------------------------------------------------
        # Step 1b: Load custom aliases (from the "Aliases" sheet, if it exists)
        # -------------------------------------------------------------------
        # Custom aliases are name variants you've encountered in real datasets
        # that don't appear in any standard source. They persist across sessions.
        self._custom_aliases = {}  # {alias_string: isocode}
        try:
            alias_df = pd.read_excel(self.filepath, sheet_name="Aliases")
            for _, row in alias_df.iterrows():
                alias = str(row["alias"]).strip()
                iso3 = str(row["isocode"]).strip().upper()
                self._custom_aliases[alias] = iso3
            log.info(f"Loaded {len(self._custom_aliases)} custom aliases from Aliases sheet")
        except (ValueError, KeyError):
            # Sheet doesn't exist yet — that's fine, it'll be created on first save
            log.info("No Aliases sheet found — will create one on first save")

        # -------------------------------------------------------------------
        # Step 1c: Build the reverse lookup index
        # -------------------------------------------------------------------
        self._index = _build_lookup_index(self.df)

        # Merge custom aliases into the index
        for alias, iso3 in self._custom_aliases.items():
            self._index[alias.lower()] = iso3

        # Also store a list of all known name variants for fuzzy matching
        self._all_names = list({
            str(v).strip()
            for col in ["country", "wb_country", "country_french"]
            for v in self.df[col].dropna()
        })

    # -----------------------------------------------------------------------
    # LOOKUP — Find one country by any key
    # -----------------------------------------------------------------------

    def lookup(self, query: str, fuzzy: bool = True) -> Optional[pd.Series]:
        """
        Look up a single country by any identifier.

        Steps:
          1. Try exact match against the index (name, ISO code, UN code, etc.)
          2. If fuzzy=True and no exact match, try fuzzy string matching
          3. Return the full row from the master DataFrame, or None

        Args:
            query: Any country name, ISO code, or numeric code
            fuzzy: Whether to attempt fuzzy matching if exact lookup fails

        Returns:
            pd.Series (the full row) if found, else None
        """
        key = str(query).strip().lower()

        # --- Exact match ---
        iso3 = self._index.get(key)
        if iso3:
            row = self.df[self.df["isocode"] == iso3]
            return row.iloc[0] if not row.empty else None

        # --- Fuzzy match ---
        if fuzzy and key:
            match, score = _find_best_match(query, self._all_names)
            if match:
                if score >= self.fuzzy_threshold:
                    log.info(f"Fuzzy match: '{query}' → '{match}' (score={score})")
                    return self.lookup(match, fuzzy=False)
                else:
                    log.warning(f"No match for '{query}' (best: '{match}', score={score})")

        return None

    # -----------------------------------------------------------------------
    # HARMONIZE — Clean an entire DataFrame column
    # -----------------------------------------------------------------------

    def harmonize(
        self,
        df: pd.DataFrame,
        country_col: str,
        target_col: str = "isocode",
        add_columns: Optional[list] = None,
        fuzzy: bool = True,
    ) -> pd.DataFrame:
        """
        Auto-standardize a DataFrame column of messy country names.

        For each unique value in `country_col`, this method:
          1. Looks it up in the master index (exact then fuzzy)
          2. Maps it to the requested target column (default: ISO3 code)
          3. Optionally joins additional master columns (region, income, etc.)
          4. Flags unmatched countries for manual review

        Args:
            df: Your input DataFrame with messy country names
            country_col: Name of the column containing country names/codes
            target_col: Which master column to map to (default: 'isocode')
            add_columns: Additional master columns to join (e.g., ['region', 'income'])
            fuzzy: Whether to use fuzzy matching for non-exact matches

        Returns:
            The input DataFrame with new columns:
              - `{target_col}`: The harmonized value
              - `_match_type`: 'exact', 'fuzzy', or 'unmatched'
              - `_match_score`: Fuzzy score (100 for exact matches)
              - Plus any `add_columns` requested
        """
        result_df = df.copy()
        unique_names = df[country_col].dropna().unique()

        # ---- Build a mapping for each unique name ----
        mapping = {}      # name → {target_col: ..., _match_type: ..., _match_score: ...}
        unmatched = []     # names that couldn't be resolved

        for name in unique_names:
            key = str(name).strip().lower()
            iso3 = self._index.get(key)

            if iso3:
                # Exact match found
                row = self.df[self.df["isocode"] == iso3].iloc[0]
                mapping[name] = {
                    target_col: row[target_col],
                    "_match_type": "exact",
                    "_match_score": 100,
                }
                # Add optional extra columns
                if add_columns:
                    for col in add_columns:
                        mapping[name][col] = row.get(col)

            elif fuzzy:
                # Try fuzzy matching
                match, score = _find_best_match(str(name), self._all_names)
                if score >= self.fuzzy_threshold:
                    row_found = self.lookup(match, fuzzy=False)
                    if row_found is not None:
                        mapping[name] = {
                            target_col: row_found[target_col],
                            "_match_type": "fuzzy",
                            "_match_score": score,
                        }
                        if add_columns:
                            for col in add_columns:
                                mapping[name][col] = row_found.get(col)
                        continue

                # Below threshold — flag for review
                mapping[name] = {
                    target_col: None,
                    "_match_type": "unmatched",
                    "_match_score": score if fuzzy else 0,
                }
                unmatched.append((name, match, score))
            else:
                mapping[name] = {
                    target_col: None,
                    "_match_type": "unmatched",
                    "_match_score": 0,
                }
                unmatched.append((name, "", 0))

        # ---- Apply mapping to DataFrame ----
        map_df = pd.DataFrame.from_dict(mapping, orient="index")
        map_df.index.name = country_col

        # Join the mapping onto the original DataFrame
        cols_to_add = [target_col, "_match_type", "_match_score"]
        if add_columns:
            cols_to_add += [c for c in add_columns if c in map_df.columns]

        for col in cols_to_add:
            result_df[col] = result_df[country_col].map(
                map_df[col].to_dict() if col in map_df.columns else {}
            )

        # ---- Report unmatched for manual review ----
        if unmatched:
            log.warning(f"\n{'='*60}")
            log.warning(f"  {len(unmatched)} UNMATCHED countries flagged for review:")
            log.warning(f"{'='*60}")
            for original, best_guess, score in unmatched:
                log.warning(f"  '{original}' → closest: '{best_guess}' (score={score})")
            log.warning(f"{'='*60}")

        matched = len(unique_names) - len(unmatched)
        log.info(f"Harmonization complete: {matched}/{len(unique_names)} matched, "
                 f"{len(unmatched)} unmatched")

        return result_df

    # -----------------------------------------------------------------------
    # UPDATE FROM WORLD BANK API
    # -----------------------------------------------------------------------

    def update_from_wb(self) -> dict:
        """
        Refresh World Bank classifications from the live API.

        Process:
          1. Fetch current WB country data via API
          2. Match to existing rows using wb_isocode
          3. Detect changes (new values vs old values)
          4. Apply updates to the in-memory DataFrame
          5. Flag new countries in WB data not in the master file

        Returns:
            dict summarizing what changed:
                {
                    "updated_cells": [(isocode, column, old_val, new_val), ...],
                    "new_wb_countries": [list of iso codes not in master],
                    "timestamp": "..."
                }
        """
        wb_df = fetch_wb_countries()

        changes = {
            "updated_cells": [],
            "new_wb_countries": [],
            "timestamp": datetime.now().isoformat(),
        }

        # Index the master file by wb_isocode for fast matching
        master_iso_set = set(self.df["isocode"].str.upper())

        for _, wb_row in wb_df.iterrows():
            wb_iso = wb_row["wb_isocode"].strip().upper()

            # Find the matching row in our master file
            mask = self.df["isocode"].str.upper() == wb_iso
            if not mask.any():
                # Also try wb_isocode column (some codes differ)
                mask = self.df["wb_isocode"].str.upper() == wb_iso
                if not mask.any():
                    changes["new_wb_countries"].append(wb_iso)
                    continue

            idx = self.df[mask].index[0]

            # Compare and update each WB column
            for col in WB_UPDATEABLE_COLS:
                old_val = self.df.at[idx, col]
                new_val = wb_row.get(col, "")

                # Normalize for comparison (handle NaN, whitespace)
                old_norm = str(old_val).strip() if pd.notna(old_val) else ""
                new_norm = str(new_val).strip() if pd.notna(new_val) else ""

                if old_norm != new_norm and new_norm:
                    self.df.at[idx, col] = new_val
                    changes["updated_cells"].append(
                        (wb_iso, col, old_norm, new_norm)
                    )

        # ---- Report results ----
        n_updates = len(changes["updated_cells"])
        n_new = len(changes["new_wb_countries"])

        if n_updates:
            log.info(f"\n{'='*60}")
            log.info(f"  {n_updates} cells updated from World Bank API:")
            log.info(f"{'='*60}")
            for iso, col, old, new in changes["updated_cells"]:
                log.info(f"  {iso}.{col}: '{old}' → '{new}'")

        if n_new:
            log.warning(f"\n  {n_new} countries in WB data not found in master file:")
            for iso in changes["new_wb_countries"]:
                log.warning(f"    {iso}")

        if not n_updates and not n_new:
            log.info("No changes detected — master file is up to date with World Bank.")

        return changes

    # -----------------------------------------------------------------------
    # UPDATE FROM IMF DATAMAPPER API
    # -----------------------------------------------------------------------

    def update_from_imf(self) -> dict:
        """
        Refresh IMF classifications (region + income) from the DataMapper API.

        The IMF DataMapper provides two things we care about:
          - Analytical groups → maps to the 'income' column
            (Advanced Economies, Emerging Market Economies, Low Income Developing Countries)
          - Regional groupings → maps to the 'region' column
            (Sub-Saharan Africa, Latin America and the Caribbean, etc.)

        Process:
          1. Fetch current classifications from the IMF DataMapper API
          2. Match to existing rows using isocode
          3. Detect and apply changes
          4. Flag countries in IMF data not in the master file

        Returns:
            dict: {"updated_cells": [...], "new_imf_countries": [...], "timestamp": "..."}
        """
        imf_data = fetch_imf_classifications()

        changes = {
            "updated_cells": [],
            "new_imf_countries": [],
            "timestamp": datetime.now().isoformat(),
        }

        master_iso_set = set(self.df["isocode"].str.upper())

        for iso3, classifications in imf_data.items():
            iso3_upper = iso3.strip().upper()

            # Find matching row in master file
            mask = self.df["isocode"].str.upper() == iso3_upper
            if not mask.any():
                changes["new_imf_countries"].append(iso3_upper)
                continue

            idx = self.df[mask].index[0]

            # Compare and update each IMF column
            for col in IMF_UPDATEABLE_COLS:
                new_val = classifications.get(col, "")
                if not new_val:
                    continue

                old_val = self.df.at[idx, col]
                old_norm = str(old_val).strip() if pd.notna(old_val) else ""
                new_norm = str(new_val).strip()

                if old_norm != new_norm:
                    self.df.at[idx, col] = new_val
                    changes["updated_cells"].append(
                        (iso3_upper, col, old_norm, new_norm)
                    )

        # ---- Report results ----
        n_updates = len(changes["updated_cells"])
        n_new = len(changes["new_imf_countries"])

        if n_updates:
            log.info(f"\n{'='*60}")
            log.info(f"  {n_updates} cells updated from IMF DataMapper API:")
            log.info(f"{'='*60}")
            for iso, col, old, new in changes["updated_cells"]:
                log.info(f"  {iso}.{col}: '{old}' → '{new}'")

        if n_new:
            log.warning(f"\n  {n_new} countries in IMF data not found in master file:")
            for iso in changes["new_imf_countries"][:20]:
                log.warning(f"    {iso}")

        if not n_updates and not n_new:
            log.info("No changes detected — master file is up to date with IMF.")

        return changes

    # -----------------------------------------------------------------------
    # ADD ALIAS — Register a new name variant manually
    # -----------------------------------------------------------------------

    def add_alias(self, alias: str, isocode: str):
        """
        Register a custom name variant so future lookups resolve it.

        Useful when you encounter a dataset with a non-standard name
        (e.g., 'Timor Leste' instead of 'Timor-Leste') and want to
        remember it permanently.

        Args:
            alias: The non-standard name to register
            isocode: The ISO3 code it should map to
        """
        key = alias.strip().lower()
        iso3 = isocode.strip().upper()

        if iso3 not in self.df["isocode"].values:
            raise ValueError(f"ISO code '{iso3}' not found in master file")

        self._index[key] = iso3
        self._custom_aliases[alias.strip()] = iso3  # Track for persistence
        log.info(f"Alias registered: '{alias}' → {iso3}")

    # -----------------------------------------------------------------------
    # SAVE — Write back to Excel with version control
    # -----------------------------------------------------------------------

    def save(self, message: str = "Update master country list"):
        """
        Save the current state back to the master Excel file.

        Version control steps:
          1. Create a timestamped backup of the current file
          2. Write the updated DataFrame to the Excel file
          3. Append an entry to the changelog
          4. If inside a Git repo, commit the change automatically

        Args:
            message: Description of what changed (used in changelog + Git commit)
        """
        # Step 1: Backup the existing file
        if self.filepath.exists():
            _create_backup(self.filepath)

        # Step 2: Write Main data + Aliases sheet (both in the same file)
        # Using ExcelWriter to write multiple sheets into one workbook
        alias_df = pd.DataFrame([
            {"alias": alias, "isocode": iso3}
            for alias, iso3 in sorted(self._custom_aliases.items())
        ])

        with pd.ExcelWriter(self.filepath, engine="openpyxl") as writer:
            self.df.to_excel(writer, sheet_name="Main", index=False)
            alias_df.to_excel(writer, sheet_name="Aliases", index=False)

        log.info(f"Master file saved: {self.filepath.name} "
                 f"({len(self._custom_aliases)} custom aliases)")

        # Step 3: Rebuild the index (data may have changed)
        self._index = _build_lookup_index(self.df)

        # Re-merge custom aliases into the freshly rebuilt index
        for alias, iso3 in self._custom_aliases.items():
            self._index[alias.lower()] = iso3

        self._all_names = list({
            str(v).strip()
            for col in ["country", "wb_country", "country_french"]
            for v in self.df[col].dropna()
        })

        # Step 4: Append to changelog
        _append_changelog(self.filepath, {
            "timestamp": datetime.now().isoformat(),
            "action": message,
            "file_hash": hashlib.md5(
                self.filepath.read_bytes()
            ).hexdigest(),
            "rows": self.df.shape[0],
            "columns": self.df.shape[1],
        })

        # Step 5: Git commit (if applicable)
        if _is_git_repo(self.filepath.parent):
            _git_commit(self.filepath, message)
        else:
            log.info("Not a Git repo — skipping auto-commit. "
                     "Backup saved to /backups/ folder instead.")

    # -----------------------------------------------------------------------
    # REPORT — Summary of the master file's current state
    # -----------------------------------------------------------------------

    def report(self) -> str:
        """Print a quick health-check summary of the master file."""
        lines = [
            f"\n{'='*60}",
            f"  Master Country List — Health Report",
            f"{'='*60}",
            f"  File:       {self.filepath.name}",
            f"  Countries:  {self.df.shape[0]}",
            f"  Columns:    {self.df.shape[1]}",
            f"  Index size: {len(self._index)} lookup entries",
            f"",
            f"  --- Completeness ---",
        ]

        for col in self.df.columns:
            n_missing = self.df[col].isna().sum()
            pct = (1 - n_missing / len(self.df)) * 100
            bar = "█" * int(pct // 5) + "░" * (20 - int(pct // 5))
            lines.append(f"  {col:25s} {bar} {pct:5.1f}% ({n_missing} missing)")

        # Membership flags summary
        flags = ["oil_export", "eu", "euro_zone", "g7", "hipc", "oecd", "prgt"]
        lines.append(f"\n  --- Membership Flags ---")
        for f in flags:
            count = self.df[f].notna().sum()
            lines.append(f"  {f:15s}: {count:3d} countries")

        lines.append(f"{'='*60}")
        report_text = "\n".join(lines)
        print(report_text)
        return report_text


# ===========================================================================
# STEP 5 — Command-Line Interface
# ===========================================================================

def main():
    """
    CLI entry point for common operations.

    Commands:
      update   — Fetch latest World Bank data and update the master file
      report   — Print a health-check summary of the master file
      lookup   — Look up a single country by any key
      harmonize — Harmonize a CSV file's country column against the master
    """
    parser = argparse.ArgumentParser(
        description="Country Harmonizer — Master Country List Manager",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Update from World Bank API
  python country_harmonizer.py update --file MasterCountryList.xlsx

  # Look up a single country
  python country_harmonizer.py lookup --file MasterCountryList.xlsx --query "Russie"

  # Health report
  python country_harmonizer.py report --file MasterCountryList.xlsx

  # Harmonize a CSV
  python country_harmonizer.py harmonize --file MasterCountryList.xlsx \\
      --input data.csv --country-col "Country Name" --output cleaned.csv
        """,
    )

    parser.add_argument("command", choices=["update", "report", "lookup", "harmonize"],
                        help="Action to perform")
    parser.add_argument("--file", "-f", required=True,
                        help="Path to the master Excel file")
    parser.add_argument("--query", "-q",
                        help="Country name/code to look up (for 'lookup' command)")
    parser.add_argument("--input", "-i",
                        help="Input CSV to harmonize (for 'harmonize' command)")
    parser.add_argument("--country-col", default="country",
                        help="Column name containing country names in the input CSV")
    parser.add_argument("--output", "-o",
                        help="Output path for harmonized CSV")
    parser.add_argument("--threshold", type=int, default=FUZZY_THRESHOLD,
                        help=f"Fuzzy match threshold (default: {FUZZY_THRESHOLD})")

    args = parser.parse_args()

    # Initialize the harmonizer
    ch = CountryHarmonizer(args.file, fuzzy_threshold=args.threshold)

    # ---- Dispatch to the right command ----

    if args.command == "report":
        ch.report()

    elif args.command == "lookup":
        if not args.query:
            parser.error("--query is required for the 'lookup' command")
        result = ch.lookup(args.query)
        if result is not None:
            print(f"\nMatch found for '{args.query}':")
            print(result.to_string())
        else:
            print(f"\nNo match found for '{args.query}'")

    elif args.command == "update":
        # Run both API updates and collect total changes
        total_cells = 0
        any_changes = False

        # --- World Bank ---
        print("Updating from World Bank API...")
        try:
            wb_changes = ch.update_from_wb()
            total_cells += len(wb_changes["updated_cells"])
            if wb_changes["updated_cells"] or wb_changes["new_wb_countries"]:
                any_changes = True
        except ConnectionError as e:
            print(f"  ⚠ World Bank API unavailable: {e}")
            wb_changes = {"updated_cells": [], "new_wb_countries": []}

        # --- IMF ---
        print("Updating from IMF DataMapper API...")
        try:
            imf_changes = ch.update_from_imf()
            total_cells += len(imf_changes["updated_cells"])
            if imf_changes["updated_cells"] or imf_changes["new_imf_countries"]:
                any_changes = True
        except ConnectionError as e:
            print(f"  ⚠ IMF API unavailable: {e}")
            imf_changes = {"updated_cells": [], "new_imf_countries": []}

        # --- Save if anything changed ---
        if any_changes:
            ch.save(message=f"API update: {total_cells} cells changed "
                    f"(WB: {len(wb_changes['updated_cells'])}, "
                    f"IMF: {len(imf_changes['updated_cells'])})")
            print(f"\nMaster file updated and saved ({total_cells} cells changed).")
        else:
            print(f"\nNo changes needed — file is up to date.")

    elif args.command == "harmonize":
        if not args.input:
            parser.error("--input is required for the 'harmonize' command")
        input_df = pd.read_csv(args.input)
        result = ch.harmonize(input_df, country_col=args.country_col)

        out_path = args.output or args.input.replace(".csv", "_harmonized.csv")
        result.to_csv(out_path, index=False)
        print(f"\nHarmonized file saved to: {out_path}")

        # Show unmatched summary
        unmatched = result[result["_match_type"] == "unmatched"]
        if not unmatched.empty:
            print(f"\n⚠ {len(unmatched)} rows unmatched — review '_match_type' column")


if __name__ == "__main__":
    main()
