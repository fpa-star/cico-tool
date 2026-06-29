"""
contra_matcher.py
Post-classification contra matching.

Logic:
1. Pivot all Contra / Contra-Unidentified rows by Value Date → sum Net.
2. Dates where sum == 0 → confirm all as Contra.
3. Dates where sum != 0 → try to find debit/credit pairs that net to zero.
4. Flag remaining unmatched entries.
"""

import logging
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

DSA_ACCOUNTS = {
    "GSPL_7445_ICICI DSA", "GWPL-ICICI-5734",
    "GSPL_1248_HDFC DSA", "GSPL_1251_HDFC DSA",
    "GSPL_3352_Yes DSA", "GSPL_3342_Yes DSA",
    "GWPL-HDFC-3572", "GWPL-Indusind-9726",
}


def _to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def run_contra_matching(df) -> dict:
    """
    Run contra matching on the classified DataFrame.
    Modifies Head in place for resolved rows.
    Returns a summary dict.
    """
    COL_ACCOUNT = "Account Name"
    COL_DATE    = "Value Date"
    COL_HEAD    = "Head"
    COL_SUBHEAD = "Sub-head 1"
    COL_SOURCE  = "Sub-head 2"
    COL_NET     = "Net"
    COL_NAR     = "Narration"
    COL_DEBIT   = "Debit"
    COL_CREDIT  = "Credit"

    contra_heads = {"Contra", "Contra - Unidentified"}

    # Work only on Contra / Contra-Unidentified rows (excluding DSA accounts)
    mask = (
        df[COL_HEAD].isin(contra_heads) &
        ~df[COL_ACCOUNT].isin(DSA_ACCOUNTS)
    )
    contra_df = df[mask].copy()

    if contra_df.empty:
        return {"dates_resolved": 0, "unmatched": [], "pairs_matched": 0}

    contra_df["_net_float"] = contra_df[COL_NET].apply(_to_float)

    resolved_count = 0
    pairs_matched  = 0
    unmatched      = []

    dates = contra_df[COL_DATE].unique()

    for date in dates:
        day_mask = mask & (df[COL_DATE] == date)
        day_rows = df[day_mask]

        net_sum = day_rows[COL_NET].apply(_to_float).sum()

        if abs(net_sum) < 0.01:  # effectively zero
            # Confirm all as Contra
            df.loc[day_mask, COL_HEAD] = "Contra"
            resolved_count += 1
        else:
            # Try pair matching within the day
            debits  = [(idx, _to_float(r[COL_DEBIT]))
                       for idx, r in day_rows.iterrows() if _to_float(r[COL_DEBIT]) > 0]
            credits = [(idx, _to_float(r[COL_CREDIT]))
                       for idx, r in day_rows.iterrows() if _to_float(r[COL_CREDIT]) > 0]

            matched_debit_idx  = set()
            matched_credit_idx = set()

            for d_idx, d_amt in debits:
                for c_idx, c_amt in credits:
                    if c_idx in matched_credit_idx:
                        continue
                    if abs(d_amt - c_amt) < 0.01:
                        df.at[d_idx, COL_HEAD] = "Contra"
                        df.at[c_idx, COL_HEAD] = "Contra"
                        matched_debit_idx.add(d_idx)
                        matched_credit_idx.add(c_idx)
                        pairs_matched += 1
                        break

            # Flag remaining unmatched
            for idx, r in day_rows.iterrows():
                if idx not in matched_debit_idx and idx not in matched_credit_idx:
                    unmatched.append({
                        "date":      str(date),
                        "account":   r.get(COL_ACCOUNT, ""),
                        "narration": r.get(COL_NAR, ""),
                        "net":       _to_float(r.get(COL_NET)),
                        "head":      r.get(COL_HEAD, ""),
                    })

    return {
        "dates_resolved":   resolved_count,
        "pairs_matched":    pairs_matched,
        "unmatched":        unmatched,
        "unmatched_count":  len(unmatched),
    }
