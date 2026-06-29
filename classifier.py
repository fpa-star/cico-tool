"""
classifier.py
Three-layer classification engine for CICO bank statement rows.

Layer 0: DSA Skip
Layer 1: Granular Lookup (classification_rules.json)
Layer 2: Pattern Rules  (pattern_rules.json + account_master.json)
Layer 3: AI Fallback    (Anthropic API)
"""

import csv
import json
import logging
import os
import time
from datetime import datetime
from typing import Optional

import anthropic

logger = logging.getLogger(__name__)

# ── Paths ─────────────────────────────────────────────────────────────────────
_BASE = os.path.dirname(__file__)
_DATA = os.path.join(_BASE, "data")
_LOGS = os.path.join(_BASE, "logs") if os.access(_BASE, os.W_OK) else "/tmp"

RULES_JSON    = os.path.join(_DATA, "classification_rules.json")
PATTERNS_JSON = os.path.join(_DATA, "pattern_rules_raw.json")
ACCOUNT_JSON  = os.path.join(_DATA, "account_master.json")
HEADS_JSON    = os.path.join(_DATA, "valid_heads.json")
AI_LOG_CSV    = os.path.join(_LOGS, "cico_ai_log.csv")

# ── DSA account names (Layer 0) ───────────────────────────────────────────────
DSA_ACCOUNTS = {
    "GSPL_7445_ICICI DSA",
    "GWPL-ICICI-5734",
    "GSPL_1248_HDFC DSA",
    "GSPL_1251_HDFC DSA",
    "GSPL_3352_Yes DSA",
    "GSPL_3342_Yes DSA",
    "GWPL-HDFC-3572",
    "GWPL-Indusind-9726",
}

# HSBC account names used in person-name detection (Layer 2)
HSBC_ACCOUNTS = {
    "GSPL-HSBC - 1001",
    "GAPL-HSBC-9001",
    "GCPL-HSBC-9001",
    "GWPL-HSBC-1001",
}

# Keywords that indicate a known company/NBFC, not a person name
HSBC_COMPANY_KEYWORDS = {
    "NORTHERN ARC", "GYANDHAN", "LORD KRISHNA", "AUXILO", "NOPAPER",
    "CREDIT CARD PAYMENT", "IN HSBC", "EROUTE", "QUATTRO", "DEPOSIT INTEREST",
    "DEBIT INTEREST", "AIRTEL", "RAB PRODUCTIONS", "SETTLEMENT", "CLOUDKEEPER",
    "TRUECALLER", "MAKEMYTRIP", "SMAS AUTO", "HARMONY", "WAVE LEAF",
    "SHREE BALAJI", "VERTIV", "KOTAK ALTERNATE", "DHIR", "ASIF MUBARAK",
    "T R CHADHA", "SOBEK AUTO", "POSITIVE ADS", "IMAGEKIT", "TLG INDIA",
    "KAS CYBER", "WITQUALIS", "M-WORTH", "ACTION X",
    "SOCIETY OF INDIAN AUTOMOBILE", "ADVANTAGE CLUB", "VLT TOURS",
    "HOST INC", "RAMPWIN",
}


def _load_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _nar(narration) -> str:
    """Return safe upper-cased narration string."""
    if narration is None:
        return ""
    try:
        return str(narration).strip().upper()
    except Exception:
        return ""


def _raw(narration) -> str:
    if narration is None:
        return ""
    try:
        return str(narration).strip()
    except Exception:
        return ""


# ── Pattern rule matching helper ──────────────────────────────────────────────

def _matches_pattern_rule(rule: dict, account_name: str, direction: str,
                          narration: str, amount: float, owning_entity: str) -> bool:
    """
    Returns True if a pattern rule matches the given row.
    direction is 'debit' or 'credit'.
    narration is already upper-cased.
    """
    r_dir = rule.get("direction", "any")
    if r_dir != "any":
        if r_dir != direction:
            return False

    # Account constraints
    if "account" in rule:
        if account_name != rule["account"]:
            return False
    if "account_in" in rule:
        if account_name not in rule["account_in"]:
            return False
    if "exclude_account" in rule:
        if account_name == rule["exclude_account"]:
            return False
    if "account_group" in rule:
        if rule["account_group"] == "HSBC":
            if account_name not in HSBC_ACCOUNTS:
                return False
    if "non_dsa_account" in rule and rule["non_dsa_account"]:
        if account_name in DSA_ACCOUNTS:
            return False

    # Narration constraints
    if "narration_contains" in rule:
        if _nar(rule["narration_contains"]) not in narration:
            return False
    if "narration_contains_any" in rule:
        if not any(_nar(kw) in narration for kw in rule["narration_contains_any"]):
            return False
    if "narration_contains_all" in rule:
        if not all(_nar(kw) in narration for kw in rule["narration_contains_all"]):
            return False
    # narration_contains_all_any: list of groups, at least one from first group AND one from second
    if "narration_contains_all_any" in rule:
        groups = rule["narration_contains_all_any"]
        for group in groups:
            if not any(_nar(kw) in narration for kw in group):
                return False

    # Amount constraints
    if "amount_gt" in rule:
        if amount <= rule["amount_gt"]:
            return False
    if "amount_lt" in rule:
        if amount >= rule["amount_lt"]:
            return False

    # Special: HSBC person name check
    if rule.get("hsbc_person_name_check"):
        if any(kw in narration for kw in HSBC_COMPANY_KEYWORDS):
            return False

    return True


def _resolve_subhead(subhead: str, owning_entity: str) -> str:
    if subhead == "{owning_entity}":
        return owning_entity or ""
    return subhead


# ── Main Classifier class ─────────────────────────────────────────────────────

class CICOClassifier:
    def __init__(self, api_key: Optional[str] = None):
        self._rules        = _load_json(RULES_JSON) if os.path.exists(RULES_JSON) else []
        self._patterns     = _load_json(PATTERNS_JSON) if os.path.exists(PATTERNS_JSON) else []
        self._account_mst  = _load_json(ACCOUNT_JSON)
        self._valid_heads  = _load_json(HEADS_JSON)
        self._valid_head_set = set(self._valid_heads["heads"])
        self._api_key      = api_key
        self._ai_client    = None
        if api_key:
            self._ai_client = anthropic.Anthropic(api_key=api_key)
        os.makedirs(_LOGS, exist_ok=True)
        self._ai_log_init()

        # Build lookup index: account_name -> list of rules (preserving order)
        self._rule_index: dict[str, list] = {}
        for rule in self._rules:
            key = rule.get("account_name", "").strip()
            self._rule_index.setdefault(key, []).append(rule)

    def reload_rules(self):
        self._rules = _load_json(RULES_JSON) if os.path.exists(RULES_JSON) else []
        self._rule_index = {}
        for rule in self._rules:
            key = rule.get("account_name", "").strip()
            self._rule_index.setdefault(key, []).append(rule)

    # ── AI log ────────────────────────────────────────────────────────────────

    def _ai_log_init(self):
        if not os.path.exists(AI_LOG_CSV):
            with open(AI_LOG_CSV, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "timestamp", "row_index", "account_name", "direction",
                    "narration", "amount", "head_assigned", "subhead_assigned", "raw_response"
                ])

    def _ai_log(self, row_index, account_name, direction, narration, amount,
                head, subhead, raw):
        with open(AI_LOG_CSV, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                datetime.now().isoformat(), row_index, account_name, direction,
                narration, amount, head, subhead, raw
            ])

    # ── Layer 0 ───────────────────────────────────────────────────────────────

    def _layer0(self, account_name: str):
        return account_name in DSA_ACCOUNTS

    # ── Layer 1 ───────────────────────────────────────────────────────────────

    def _layer1(self, account_name: str, direction: str, narration_raw: str):
        """
        Returns (head, subhead) or (None, None) if no rule matches.
        direction: 'debit' | 'credit'
        """
        candidate_rules = self._rule_index.get(account_name, [])
        nar_upper = _nar(narration_raw)

        # Sort: "All" rules first within the account rules
        all_rules  = [r for r in candidate_rules if r.get("rule_feature", "").strip().lower() == "all"]
        text_rules = [r for r in candidate_rules if r.get("rule_feature", "").strip().lower() != "all"]

        for rule in all_rules + text_rules:
            r_dir = rule.get("direction", "").strip().lower()
            if r_dir and r_dir not in ("", "any") and r_dir != direction:
                continue

            feature = rule.get("rule_feature", "").strip().lower()
            text    = rule.get("text", "").strip()

            matched = False
            if feature == "all":
                matched = True
            elif feature == "contains":
                matched = text.upper() in nar_upper
            elif feature in ("starts with", "startswith", "starts_with"):
                matched = nar_upper.startswith(text.upper())
            else:
                # Unknown feature — treat as contains
                matched = text.upper() in nar_upper if text else True

            if matched:
                return rule.get("head", ""), rule.get("subhead", "")

        return None, None

    # ── Layer 2 ───────────────────────────────────────────────────────────────

    def _layer2(self, account_name: str, direction: str, narration_raw: str,
                amount: float, owning_entity: str):
        """
        Apply pattern rules (priority 1-10) then account master fallback.
        Returns (head, subhead) or (None, None).
        """
        nar_upper = _nar(narration_raw)

        # Sort pattern groups by priority
        sorted_groups = sorted(self._patterns, key=lambda g: g.get("priority", 99))

        for group in sorted_groups:
            for rule in group.get("rules", []):
                if _matches_pattern_rule(rule, account_name, direction,
                                         nar_upper, amount, owning_entity):
                    raw_subhead = rule.get("subhead", "")
                    subhead = _resolve_subhead(raw_subhead, owning_entity)
                    return rule.get("head", ""), subhead

        # Account master fallback (Priority 11)
        master = self._account_mst.get(account_name)
        if master:
            entry = master.get(direction, {})
            head = entry.get("head", "")
            subhead = entry.get("subhead", "")
            if head and head != "narration":
                return head, subhead

        return None, None

    # ── Layer 3 (AI) ─────────────────────────────────────────────────────────

    def _build_ai_system_prompt(self) -> str:
        heads_str = "\n".join(f"  - {h}" for h in self._valid_heads["heads"])
        acct_str  = json.dumps(self._account_mst, indent=2)
        return f"""You are a financial transaction classifier for Girnar Software's FP&A team.
Your task: classify each bank transaction into exactly one Head and Sub-head.

VALID HEAD NAMES (use EXACTLY as written, including spaces and punctuation):
{heads_str}

ACCOUNT MASTER (account defaults):
{acct_str}

RULES:
1. Return ONLY valid JSON: {{"head": "...", "subhead": "..."}}
2. Use exact head names from the list above — no variations.
3. If uncertain between two heads, prefer the more specific one.
4. Contra = internal transfers between Girnar group accounts.
5. Sales Collection = money received from customers / dealerships.
6. Treasury receipts = interest income on FDs.
7. Sale / Purchase of FD & Investments = FD creation or redemption.
8. For salary accounts, Debit is almost always Salary and Related.
9. DSA accounts should never reach you — if they do, return {{"head": "Contra - DSA accounts", "subhead": ""}}.
"""

    def _ai_classify_batch(self, rows: list, row_indices: list) -> list:
        """
        Classify a batch of rows using the AI.
        rows: list of dicts with keys: account_name, direction, narration, amount
        Returns list of (head, subhead) tuples.
        """
        if not self._ai_client:
            return [("TBD", "") for _ in rows]

        system_prompt = self._build_ai_system_prompt()
        messages = []
        for i, row in enumerate(rows):
            messages.append(
                f"[{i}] Account: {row['account_name']}, "
                f"Direction: {row['direction']}, "
                f"Narration: {row['narration']}, "
                f"Amount: {row['amount']}"
            )

        user_msg = (
            "Classify each transaction. Respond with a JSON array, one object per transaction "
            "in the same order. Each object: {\"head\": \"...\", \"subhead\": \"...\"}.\n\n"
            + "\n".join(messages)
        )

        try:
            response = self._ai_client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=2048,
                system=system_prompt,
                messages=[{"role": "user", "content": user_msg}],
            )
            raw = response.content[0].text.strip()

            # Extract JSON array from response
            start = raw.find("[")
            end   = raw.rfind("]") + 1
            if start == -1 or end == 0:
                # Try single object
                start = raw.find("{")
                end   = raw.rfind("}") + 1
                if start != -1 and end > 0:
                    obj = json.loads(raw[start:end])
                    results = [obj] * len(rows)
                else:
                    results = [{"head": "TBD", "subhead": ""}] * len(rows)
            else:
                results = json.loads(raw[start:end])

            output = []
            for i, (result, row, ridx) in enumerate(zip(results, rows, row_indices)):
                head    = result.get("head", "TBD")
                subhead = result.get("subhead", "")
                # Validate head
                if head not in self._valid_head_set:
                    head = "TBD"
                self._ai_log(ridx, row["account_name"], row["direction"],
                             row["narration"], row["amount"], head, subhead, raw)
                output.append((head, subhead))
            return output

        except Exception as exc:
            logger.error(f"AI batch classification failed: {exc}")
            return [("TBD", "") for _ in rows]

    # ── Main classify_row ──────────────────────────────────────────────────────

    def classify_row(self, row_index: int, account_name: str, direction: str,
                     narration: str, amount: float, owning_entity: str):
        """
        Classify a single row through Layers 0-2.
        Returns (head, subhead, source) where source is 'DSA-SKIP'|'RULE'|'RULE'|'TBD'.
        Layer 3 (AI) is handled in batch via classify_dataframe.
        """
        account_name  = (account_name or "").strip()
        direction     = (direction or "").strip().lower()
        owning_entity = (owning_entity or "").strip()

        # Layer 0
        if self._layer0(account_name):
            return "", "", "DSA-SKIP"

        # Layer 1
        head, subhead = self._layer1(account_name, direction, narration)
        if head:
            return head, subhead, "RULE"

        # Layer 2
        head, subhead = self._layer2(account_name, direction, narration, amount, owning_entity)
        if head and head != "narration":
            return head, subhead, "RULE"

        return "TBD", "", "TBD"

    # ── Batch classify entire DataFrame ───────────────────────────────────────

    def classify_dataframe(self, df, progress_callback=None):
        """
        Classify all rows in a pandas DataFrame.
        Expected columns: Account Name, Debit/Credit (or derived 'direction'),
                          Narration, Debit, Credit, Net, Owning Entity.
        Adds/updates columns: Head (O), Sub-head 1 (P), ClassSource (Q).
        Returns (df, stats_dict).
        """
        import pandas as pd

        COL_ACCOUNT   = "Account Name"
        COL_DIRECTION = "Transaction Type"   # typically "DR"/"CR"; fallback computed from Debit/Credit
        COL_NARRATION = "Narration"
        COL_DEBIT     = "Debit"
        COL_CREDIT    = "Credit"
        COL_NET       = "Net"
        COL_OWNING    = "Owning Entity"
        COL_HEAD      = "Head"
        COL_SUBHEAD   = "Sub-head 1"
        COL_SOURCE    = "Sub-head 2"          # col Q — classification source

        stats = {
            "total": 0, "dsa_skip": 0, "layer1": 0,
            "layer2": 0, "ai": 0, "tbd": 0, "errors": [],
        }

        # Initialise output columns if missing
        for col in [COL_HEAD, COL_SUBHEAD, COL_SOURCE]:
            if col not in df.columns:
                df[col] = ""

        tbd_indices  = []   # rows needing AI classification
        tbd_rows     = []

        total = len(df)
        stats["total"] = total

        for i, (df_idx, row) in enumerate(df.iterrows()):
            try:
                account_name  = str(row.get(COL_ACCOUNT) or "").strip()
                narration     = str(row.get(COL_NARRATION) or "").strip()
                owning_entity = str(row.get(COL_OWNING) or "").strip()

                # Determine direction
                debit_val  = _to_float(row.get(COL_DEBIT))
                credit_val = _to_float(row.get(COL_CREDIT))
                txn_type   = str(row.get(COL_DIRECTION) or "").strip().upper()

                if txn_type in ("DR", "DEBIT"):
                    direction = "debit"
                elif txn_type in ("CR", "CREDIT"):
                    direction = "credit"
                elif debit_val and debit_val > 0:
                    direction = "debit"
                elif credit_val and credit_val > 0:
                    direction = "credit"
                else:
                    direction = "debit"  # fallback

                amount = debit_val if direction == "debit" else credit_val

                head, subhead, source = self.classify_row(
                    df_idx, account_name, direction, narration, amount, owning_entity
                )

                df.at[df_idx, COL_HEAD]    = head
                df.at[df_idx, COL_SUBHEAD] = subhead
                df.at[df_idx, COL_SOURCE]  = source

                if source == "DSA-SKIP":
                    stats["dsa_skip"] += 1
                elif source == "RULE":
                    stats["layer1"] += 1   # layers 1 & 2 both flagged RULE
                elif source == "TBD":
                    tbd_indices.append(df_idx)
                    tbd_rows.append({
                        "account_name": account_name,
                        "direction":    direction,
                        "narration":    narration,
                        "amount":       amount,
                    })

            except Exception as exc:
                stats["errors"].append(f"Row {i}: {exc}")
                df.at[df_idx, COL_SOURCE] = "TBD"

            if progress_callback and (i + 1) % 50 == 0:
                progress_callback(i + 1, total, stats)

        # Layer 3: AI batch
        if tbd_rows:
            BATCH = 50
            for b_start in range(0, len(tbd_rows), BATCH):
                batch_rows    = tbd_rows[b_start : b_start + BATCH]
                batch_indices = tbd_indices[b_start : b_start + BATCH]
                results = self._ai_classify_batch(batch_rows, batch_indices)
                for df_idx, (head, subhead) in zip(batch_indices, results):
                    df.at[df_idx, COL_HEAD]    = head
                    df.at[df_idx, COL_SUBHEAD] = subhead
                    src = "AI" if head not in ("TBD", "") else "TBD"
                    df.at[df_idx, COL_SOURCE]  = src
                    if src == "AI":
                        stats["ai"] += 1
                    else:
                        stats["tbd"] += 1
                if b_start + BATCH < len(tbd_rows):
                    time.sleep(0.5)   # gentle rate-limit buffer

        # Final counts
        if progress_callback:
            progress_callback(total, total, stats)

        return df, stats


def _to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    try:
        s = str(val).replace(",", "").strip()
        return float(s) if s else 0.0
    except (ValueError, TypeError):
        return 0.0
