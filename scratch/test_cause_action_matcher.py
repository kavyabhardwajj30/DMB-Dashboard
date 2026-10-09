import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import get_function_rca_data
import re

rca_data = get_function_rca_data()
causes = rca_data['causes']
actions = rca_data['actions']

stopwords = {'in', 'for', 'the', 'to', 'due', 'of', 'and', 'a', 'on', 'by', 'is', 'at', 'with', 'from', 'as', 'q1', 'q2', 'q3', 'q4', 'oit', 'npi', 'npis', 'kpi', 'gap', 'rate'}

def tokenize(text):
    if not text or pd.isna(text):
        return set()
    words = re.findall(r'[a-zA-Z0-9]+', str(text).lower())
    tokens = set()
    for w in words:
        if len(w) >= 3 and w not in stopwords:
            tokens.add(w)
            if w.endswith('s') and len(w) > 3:
                tokens.add(w[:-1])
            if w.endswith('ing') and len(w) > 4:
                tokens.add(w[:-3])
    return tokens

def compute_cause_action_match_score(cause_text, action_root_cause, action_desc=""):
    c_clean = str(cause_text).strip().lower()
    rc_clean = str(action_root_cause).strip().lower()
    desc_clean = str(action_desc).strip().lower()

    if not c_clean or c_clean in {"none", "nan", "—", ""}:
        return 0.0

    # 1. Exact match
    if c_clean == rc_clean:
        return 1.0

    # 2. Substring match
    if len(c_clean) >= 4 and (c_clean in rc_clean or rc_clean in c_clean):
        return 0.95

    # 3. Token overlap with root_cause (primary) and action_desc (secondary)
    c_toks = tokenize(c_clean)
    if not c_toks:
        return 0.0

    rc_toks = tokenize(rc_clean)
    desc_toks = tokenize(desc_clean)

    common_rc = c_toks & rc_toks
    common_desc = c_toks & desc_toks

    score = 0.0
    if common_rc:
        score += len(common_rc) * 0.4
    if common_desc:
        score += len(common_desc) * 0.2

    # Normalize by number of tokens in cause
    ratio = score / len(c_toks)
    return min(1.0, ratio)

print("=== TESTING CAUSE TO ACTION MATCHING ACROSS ALL KPIS ===")
for fn_key in causes['function_key'].unique():
    c_fn = causes[causes['function_key'] == fn_key]
    a_fn = actions[actions['function_key'] == fn_key]
    for kpi in c_fn['kpi_name'].unique():
        c_kpi = c_fn[c_fn['kpi_name'] == kpi]
        a_kpi = a_fn[a_fn['kpi_name'] == kpi]
        if not a_kpi.empty:
            print(f"\n[{fn_key}] KPI: {kpi}")
            for _, cr in c_kpi.iterrows():
                c_text = cr['cause']
                best_score = 0.0
                best_action = None
                for a_idx, ar in a_kpi.iterrows():
                    sc = compute_cause_action_match_score(c_text, ar['root_cause'], ar['corrective_action'])
                    if sc > best_score:
                        best_score = sc
                        best_action = ar
                
                if best_score >= 0.2:
                    print(f"  MATCHED (score={best_score:.2f}): Cause: \"{c_text}\" <==> Action RC: \"{best_action['root_cause']}\" | Desc: \"{best_action['corrective_action']}\"")
                else:
                    print(f"  NO ACTION MATCH: Cause: \"{c_text}\"")
