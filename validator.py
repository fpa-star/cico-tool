"""
validator.py
Post-classification gotcha validation and force-fix passes.
"""

import json
import os
import logging

logger = logging.getLogger(__name__)

_DATA = os.path.join(os.path.dirname(__file__), "data")
HEADS_JSON = os.path.join(_DATA, "valid_heads.json")

HSBC_ACCOUNTS = {
    "GSPL-HSBC - 1001",
    "GAPL-HSBC-9001",
    "GCPL-HSBC-9001",
    "GWPL-HSBC-1001",
}

HSBC_COMPANY_KEYWORDS = {
    "NORTHERN ARC", "GYANDHAN", "LORD KRISHNA", "AUXILO", "NOPAPER",
    "CREDIT CARD PAYMENT", "IN HSBC", "EROUTE", "QUATTRO",
    "DEPOSIT INTEREST", "DEBIT INTEREST", "AIRTEL", "RAB PRODUCTIONS",
    "SETTLEMENT", "CLOUDKEEPER", "TRUECALLER", "MAKEMYTRIP", "SMAS AUTO",
    "HARMONY", "WAVE LEAF", "SHREE BALAJI", "VERTIV", "KOTAK ALTERNATE",
    "DHIR", "ASIF MUBARAK", "T R CHADHA", "SOBEK AUTO", "POSITIVE ADS",
    "IMAGEKIT", "TLG INDIA", "KAS CYBER", "WITQUALIS", "M-WORTH",
    "ACTION X", "SOCIETY OF INDIAN AUTOMOBILE", "ADVANTAGE CLUB",
    "VLT TOURS", "HOST INC", "RAMPWIN",
}

FD_KEYWORDS = {
    "TRF TO FD", "FD CLOS", "PRIN AND INT", "FD REDEEM PRINCIPAL",
    "FD REDEEM INTEREST", "INT ON FD/RD", "NIPPON INDIA OVERNIGHT",
    "ADITYA BSLMF", "BSLMF", "FD NO.", "CLOSURE PROCEEDSCREDIT TO REPAYMENT",
}

SALARY_ACCOUNTS = {
    "GSPL-HDFC-5746 Salary",
    "GAPL-HDFC-5759",
    "GSSPL-HDFC-5851",
}


def _load_valid_heads():
    with open(HEADS_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)
    return set(data["heads"])


def _nar(val) -> str:
    if val is None:
        return ""
    return str(val).strip().upper()


def _to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def run_gotcha_checks(df) -> dict:
    """
    Run all gotcha validations on a classified DataFrame.
    Returns a dict with keys: fixes (list), flags (list).
    Each entry: {row_index, gotcha, account, narration, old_head, new_head, action}
    """
    COL_ACCOUNT  = "Account Name"
    COL_NAR      = "Narration"
    COL_HEAD     = "Head"
    COL_SUBHEAD  = "Sub-head 1"
    COL_SOURCE   = "Sub-head 2"
    COL_DEBIT    = "Debit"
    COL_CREDIT   = "Credit"

    valid_heads = _load_valid_heads()
    fixes = []
    flags = []

    for df_idx, row in df.iterrows():
        account = str(row.get(COL_ACCOUNT) or "").strip()
        nar     = _nar(row.get(COL_NAR))
        head    = str(row.get(COL_HEAD) or "").strip()
        subhead = str(row.get(COL_SUBHEAD) or "").strip()
        source  = str(row.get(COL_SOURCE) or "").strip()
        debit   = _to_float(row.get(COL_DEBIT))
        credit  = _to_float(row.get(COL_CREDIT))
        direction = "debit" if debit > 0 else "credit"
        amount    = debit if direction == "debit" else credit

        # G1: HSBC + Debit + not salary/reimbursement and not a known company
        if (account in HSBC_ACCOUNTS and direction == "debit"
                and head not in ("Salary and Related", "Interest & WCTL fees",
                                  "Other Payments", "Contra", "Sale / Purchase of FD & Investments")):
            if not any(kw in nar for kw in HSBC_COMPANY_KEYWORDS):
                flags.append({
                    "row_index": df_idx, "gotcha": "G1",
                    "account": account, "narration": nar,
                    "old_head": head, "new_head": "Salary and Related",
                    "action": "FLAG - likely Salary/Reimbursement",
                })

        # G3: GSPL-ICICI-6543 / GSPL-ICICI-3418 + Debit + "TRANSFER TO GIRNAR" → Contra
        if (account in ("GSPL-ICICI-6543", "GSPL-ICICI-3418")
                and direction == "debit"
                and "TRANSFER TO GIRNAR" in nar):
            if head != "Contra":
                df.at[df_idx, COL_HEAD]    = "Contra"
                df.at[df_idx, COL_SUBHEAD] = "GSPL"
                df.at[df_idx, COL_SOURCE]  = source + "+G3"
                fixes.append({
                    "row_index": df_idx, "gotcha": "G3",
                    "account": account, "narration": nar,
                    "old_head": head, "new_head": "Contra",
                    "action": "FIXED",
                })

        # G4: FD/investment keywords → must be Sale / Purchase of FD & Investments
        has_fd_keyword = any(kw in nar for kw in FD_KEYWORDS)
        if has_fd_keyword and head != "Sale / Purchase of FD & Investments":
            new_subhead = "FD redeem" if direction == "credit" else "FD made"
            df.at[df_idx, COL_HEAD]    = "Sale / Purchase of FD & Investments"
            df.at[df_idx, COL_SUBHEAD] = new_subhead
            df.at[df_idx, COL_SOURCE]  = source + "+G4"
            fixes.append({
                "row_index": df_idx, "gotcha": "G4",
                "account": account, "narration": nar,
                "old_head": head, "new_head": "Sale / Purchase of FD & Investments",
                "action": "FIXED",
            })

        # G7: Tax Payments / GST Paid — small amounts should be Bank charges
        if head == "Tax Payments (GST & TDS Excl. Salaries)" and subhead == "GST Paid":
            if 0 < amount < 10000:
                flags.append({
                    "row_index": df_idx, "gotcha": "G7",
                    "account": account, "narration": nar,
                    "old_head": head, "new_head": "Other Payments / Bank charges",
                    "action": "FLAG - small GST amount, check if Bank charges",
                })

        # G8: GSPL_3189_Yes Credit → should be Vendor Payment/Commission payouts
        if account == "GSPL_3189_Yes" and direction == "credit":
            if head not in ("Vendor Payment", "Contra"):
                flags.append({
                    "row_index": df_idx, "gotcha": "G8",
                    "account": account, "narration": nar,
                    "old_head": head, "new_head": "Vendor Payment / Commission payouts",
                    "action": "FLAG",
                })

        # G9: Salary account Credits
        if account in SALARY_ACCOUNTS and direction == "credit":
            if "GIRNAR" in nar and head != "Contra":
                df.at[df_idx, COL_HEAD]    = "Contra"
                df.at[df_idx, COL_SUBHEAD] = ""
                df.at[df_idx, COL_SOURCE]  = source + "+G9"
                fixes.append({
                    "row_index": df_idx, "gotcha": "G9",
                    "account": account, "narration": nar,
                    "old_head": head, "new_head": "Contra",
                    "action": "FIXED",
                })

        # G10: GIRNAR HDFC 5846 Debit → must be Sales Collection/FS-MP
        if account == "GIRNAR HDFC 5846 Rupyy Collection" and direction == "debit":
            if head != "Sales Collection":
                df.at[df_idx, COL_HEAD]    = "Sales Collection"
                df.at[df_idx, COL_SUBHEAD] = "FS-MP"
                df.at[df_idx, COL_SOURCE]  = source + "+G10"
                fixes.append({
                    "row_index": df_idx, "gotcha": "G10",
                    "account": account, "narration": nar,
                    "old_head": head, "new_head": "Sales Collection",
                    "action": "FIXED",
                })

        # G11: GAPL-HSBC-9001 Debit with C2D keywords → Sales Collection/Crack-Ed
        if account == "GAPL-HSBC-9001" and direction == "debit":
            c2d_keywords = [
                "NORTHERN ARC CAPITAL", "GYANDHAN FINANCIAL",
                "LORD KRISHNA FINANCIAL", "AKSHAR FEE MANAGEMENT",
            ]
            if any(kw in nar for kw in c2d_keywords):
                if head != "Sales Collection":
                    df.at[df_idx, COL_HEAD]    = "Sales Collection"
                    df.at[df_idx, COL_SUBHEAD] = "Crack-Ed"
                    df.at[df_idx, COL_SOURCE]  = source + "+G11"
                    fixes.append({
                        "row_index": df_idx, "gotcha": "G11",
                        "account": account, "narration": nar,
                        "old_head": head, "new_head": "Sales Collection",
                        "action": "FIXED",
                    })

        # G14: Validate head name
        re_head = str(df.at[df_idx, COL_HEAD] or "").strip()
        if re_head and re_head not in ("TBD", "", "DSA-SKIP") and re_head not in valid_heads:
            flags.append({
                "row_index": df_idx, "gotcha": "G14",
                "account": account, "narration": nar,
                "old_head": re_head, "new_head": "TBD",
                "action": f"FLAG - invalid head name: '{re_head}'",
            })

    return {"fixes": fixes, "flags": flags}
