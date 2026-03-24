# -*- coding: utf-8 -*-
"""
reporting/final_report.py
─────────────────────────
Final Client Report Generation — post-optimization workflow.

Public API
──────────
    run_planning_ai(structured_input)  -> dict[section_name, text]
    render_final_report_ui(rows_list, recs, baseline, product_type)
        → renders the full UI block inside a Streamlit expander

AI Integration
──────────────
- mode = "planning" (as spec'd)
- Prompt sourced from Google Docs (same URL as ISA module)
- Structured JSON input only
- Returns exactly 6 sections
"""
from __future__ import annotations

import json
import math
import re
import os
from typing import Optional

# ── AI helpers (reuse from ai_analyst) ───────────────────────────────────────

_GUIDANCE_DOC_URL = (
    "https://docs.google.com/document/d/"
    "1Hqh9TI2u7QRbTvRAS0TRkyL-28-eLIznkQbGTd_M1Wk"
)
_NaN = float("nan")

SECTION_KEYS = [
    "executive_summary",
    "current_weaknesses",
    "planning_principles",
    "change_advantages",
    "risks_considerations",
    "final_summary",
]

SECTION_LABELS_HE = {
    "executive_summary":     "1. תקציר מנהלים",
    "current_weaknesses":    "2. חולשות התיק הנוכחי",
    "planning_principles":   "3. עקרונות התכנון",
    "change_advantages":     "4. יתרונות השינויים המוצעים",
    "risks_considerations":  "5. שיקולים ואיזונים",
    "final_summary":         "6. סיכום סופי",
}


def _get_api_key() -> str:
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "OPENAI_API_KEY" in st.secrets:
            return str(st.secrets["OPENAI_API_KEY"])
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY", "")


def _fetch_guidance() -> str:
    """Fetch AI writing instructions from the canonical Google Doc."""
    import requests
    doc_id = re.search(r"/d/([^/]+)", _GUIDANCE_DOC_URL)
    if not doc_id:
        return ""
    did = doc_id.group(1)
    for url in [
        f"https://docs.google.com/document/d/{did}/export?format=txt",
        f"https://docs.google.com/document/u/0/d/{did}/export?format=txt",
    ]:
        try:
            r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
            if r.status_code == 200:
                txt = r.text.strip()
                if txt and "JavaScript" not in txt[:300]:
                    return txt
        except Exception:
            pass
    return ""


def _fmt(v, pct=True) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "לא זמין"
    if pct:
        return f"{v:.1f}%"
    return f"{v:.2f}"


def _build_planning_prompt(structured: dict, guidance: str) -> str:
    """Build the planning-mode AI prompt from structured JSON input."""
    pb   = structured.get("portfolio_before", {})
    pa   = structured.get("portfolio_after", {})
    obj  = structured.get("client_objectives", {})
    sol  = structured.get("selected_solution_name", "")
    chg  = structured.get("changes_summary", {})

    before_block = (
        f"  מניות: {_fmt(pb.get('equities'))}\n"
        f"  חו\"ל: {_fmt(pb.get('abroad'))}\n"
        f"  מט\"ח: {_fmt(pb.get('fx'))}\n"
        f"  לא-סחיר: {_fmt(pb.get('illiquid'))}\n"
        f"  שארפ: {_fmt(pb.get('sharpe'), pct=False)}\n"
        f"  עלות שירות: {_fmt(pb.get('cost'))}\n"
        f"  מנהלים: {pb.get('managers_count', 'לא זמין')}\n"
        f"  מוצרים: {pb.get('products_count', 'לא זמין')}"
    )
    after_block = (
        f"  מניות: {_fmt(pa.get('equities'))}\n"
        f"  חו\"ל: {_fmt(pa.get('abroad'))}\n"
        f"  מט\"ח: {_fmt(pa.get('fx'))}\n"
        f"  לא-סחיר: {_fmt(pa.get('illiquid'))}\n"
        f"  שארפ: {_fmt(pa.get('sharpe'), pct=False)}\n"
        f"  עלות שירות: {_fmt(pa.get('cost'))}\n"
        f"  מנהלים: {pa.get('managers_count', 'לא זמין')}\n"
        f"  מוצרים: {pa.get('products_count', 'לא זמין')}"
    )
    obj_block = (
        f"  יעד מניות: {_fmt(obj.get('target_equities'))}\n"
        f"  יעד חו\"ל: {_fmt(obj.get('target_abroad'))}\n"
        f"  יעד מט\"ח: {_fmt(obj.get('target_fx'))}\n"
        f"  יעד לא-סחיר: {_fmt(obj.get('target_illiquid'))}\n"
        f"  עדיפות: {obj.get('primary_rank', 'דיוק')}\n"
        f"  עולם מוצר: {obj.get('product_type', 'לא צוין')}"
    )
    delta_block = "\n".join(
        f"  {k}: {_fmt(v, pct=True)}" for k, v in chg.items()
    ) or "  (אין נתוני דלתא)"

    guidance_section = (
        f"הנחיות כתיבה ממסמך חיצוני:\n{guidance}"
        if guidance
        else "לא נטענו הנחיות חיצוניות. היצמד לנתונים בלבד."
    )

    return f"""mode: planning
נושא: הפקת דוח לקוח לאחר אופטימיזציית תמהיל.

נתוני הקלט הם JSON מובנה בלבד. אל תסיק מעבר לנתונים שנמסרו לך.

--- תיק נוכחי (לפני) ---
{before_block}

--- תיק מוצע (אחרי) ---
{after_block}

--- יעדי הלקוח ---
{obj_block}

--- שם החלופה שנבחרה ---
{sol or "לא צוין"}

--- שינויים מרכזיים (דלתא) ---
{delta_block}

---
{guidance_section}

---
הוראות תפוקה — חובה לפי הסדר הזה בדיוק:

כתוב בעברית. כל סעיף מתחיל בכותרת מפורשת:
[1. תקציר מנהלים]
[2. חולשות התיק הנוכחי]
[3. עקרונות התכנון]
[4. יתרונות השינויים המוצעים]
[5. שיקולים ואיזונים]
[6. סיכום סופי]

דרישות:
- אל תמציא נתונים שאינם בקלט
- אל תיתן ייעוץ השקעות מחייב
- הבהר את השיפור בפיזור הסיכון
- התייחס לחשיפה גלובלית, ריכוז ויעדי הלקוח
- טון: מקצועי, מאוזן, מבוסס-דאטה
- אל תשתמש בנוסח ודאות מוחלטת"""


def _parse_sections(text: str) -> dict[str, str]:
    """Parse 6 bracketed sections from the AI response."""
    patterns = {
        "executive_summary":    r"\[1\.[^\]]*\]",
        "current_weaknesses":   r"\[2\.[^\]]*\]",
        "planning_principles":  r"\[3\.[^\]]*\]",
        "change_advantages":    r"\[4\.[^\]]*\]",
        "risks_considerations": r"\[5\.[^\]]*\]",
        "final_summary":        r"\[6\.[^\]]*\]",
    }
    result = {k: "" for k in SECTION_KEYS}
    positions = []
    for key, pat in patterns.items():
        m = re.search(pat, text)
        if m:
            positions.append((m.start(), key, m.end()))

    positions.sort()
    for i, (start, key, end) in enumerate(positions):
        next_start = positions[i + 1][0] if i + 1 < len(positions) else len(text)
        result[key] = text[end:next_start].strip()

    # Fallback: if parse fails, put everything in executive_summary
    if all(v == "" for v in result.values()):
        result["executive_summary"] = text.strip()
    return result


def run_planning_ai(structured_input: dict) -> tuple[dict[str, str], Optional[str]]:
    """
    Call AI in planning mode with structured JSON input.
    Returns (sections_dict, error_string_or_None).
    """
    try:
        from institutional_strategy_analysis.ai_analyst import _call_claude
    except ImportError:
        # Fallback: direct HTTP call using the same pattern
        def _call_claude(prompt, system="", max_tokens=4000, model="gpt-4o"):
            import requests as _r
            key = _get_api_key()
            if not key:
                return "", "לא הוגדר מפתח OPENAI_API_KEY."
            msgs = []
            if system:
                msgs.append({"role": "system", "content": system})
            msgs.append({"role": "user", "content": prompt})
            try:
                resp = _r.post(
                    "https://api.openai.com/v1/chat/completions",
                    headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                    json={"model": model, "max_tokens": max_tokens, "messages": msgs},
                    timeout=90,
                )
                if resp.status_code == 200:
                    t = resp.json().get("choices", [{}])[0].get("message", {}).get("content", "").strip()
                    return (t, None) if t else ("", "תגובה ריקה.")
                return "", f"שגיאת API: HTTP {resp.status_code}"
            except Exception as e:
                return "", f"שגיאת תקשורת: {e}"

    guidance = _fetch_guidance()
    prompt   = _build_planning_prompt(structured_input, guidance)
    system   = (
        "אתה יועץ השקעות המנתח תיקי השקעות בעברית. "
        "עבד במצב planning — הפקת דוח לקוח לאחר אופטימיזציה. "
        "היצמד לנתונים שנמסרו, אל תסיק מעבר למה שנמסר, "
        "ואל תיתן ייעוץ השקעות מחייב."
    )
    raw, err = _call_claude(prompt, system=system, max_tokens=4000)
    if err:
        return {k: "" for k in SECTION_KEYS}, err
    return _parse_sections(raw), None


def build_notebook_package(
    structured_input: dict,
    sections: dict[str, str],
    product_type: str = "",
) -> str:
    """Build a Notebook-ready JSON package (as string) for export."""
    slides = [
        {"slide": 1, "title": "שער",
         "content": f"דוח לקוח — {product_type or 'אופטימיזציית תמהיל'}\nהופק על-ידי Profit Financial Group"},
        {"slide": 2, "title": "תקציר מנהלים",
         "content": sections.get("executive_summary", "")},
        {"slide": 3, "title": "תיק נוכחי",
         "content": json.dumps(structured_input.get("portfolio_before", {}), ensure_ascii=False, indent=2)},
        {"slide": 4, "title": "תיק מוצע",
         "content": json.dumps(structured_input.get("portfolio_after", {}), ensure_ascii=False, indent=2)},
        {"slide": 5, "title": "השוואה לפני / אחרי",
         "content": json.dumps(structured_input.get("changes_summary", {}), ensure_ascii=False, indent=2)},
        {"slide": 6, "title": "יתרונות השינויים",
         "content": sections.get("change_advantages", "")},
        {"slide": 7, "title": "שיקולים ואיזונים",
         "content": sections.get("risks_considerations", "")},
        {"slide": 8, "title": "סיכום סופי",
         "content": sections.get("final_summary", "")},
    ]
    pkg = {
        "mode": "planning",
        "product_type": product_type,
        "selected_solution": structured_input.get("selected_solution_name", ""),
        "client_objectives": structured_input.get("client_objectives", {}),
        "portfolio_before":  structured_input.get("portfolio_before", {}),
        "portfolio_after":   structured_input.get("portfolio_after", {}),
        "changes_summary":   structured_input.get("changes_summary", {}),
        "ai_sections": sections,
        "presentation_slides": slides,
        "presentation_prompt": (
            "צור מצגת מקצועית בעברית הכוללת 8 שקופיות לפי המבנה שלמעלה. "
            "השתמש בנתונים מ portfolio_before ו-portfolio_after לשקופיות 3-5. "
            "השתמש בטקסט מ ai_sections לשקופיות 2, 6, 7, 8. "
            "שמור על טון מקצועי ומאוזן."
        ),
    }
    return json.dumps(pkg, ensure_ascii=False, indent=2)


# ── Streamlit UI ──────────────────────────────────────────────────────────────

def render_final_report_ui(
    rows_list: list,
    recs: dict,
    baseline: Optional[dict],
    product_type: str,
) -> None:
    """
    Renders the Final Client Report expander section.
    Call this immediately after render_results_table() in streamlit_app.py.
    """
    import streamlit as st

    _nan = float("nan")

    def _f(v) -> float:
        try:
            f = float(v)
            return _nan if math.isnan(f) else f
        except (TypeError, ValueError):
            return _nan


    # ── Step 1: Build structured input ────────────────────────────
    bl   = baseline or {}
    best = (recs.get("weighted") or recs.get("accurate") or
            recs.get("sharpe")  or recs.get("service") or {})

    # Portfolio before
    pb = {
        "equities":       _f(bl.get("stocks")),
        "abroad":         _f(bl.get("foreign")),
        "fx":             _f(bl.get("fx")),
        "illiquid":       _f(bl.get("illiquid")),
        "sharpe":         _f(bl.get("sharpe")),
        "cost":           _f(bl.get("service")),
        "managers_count": int(bl.get("managers_count", 0)) if bl else 0,
        "products_count": int(bl.get("products_count", 0)) if bl else 0,
    }

    # Portfolio after
    pa = {
        "equities": _f(best.get("מניות (%)")),
        "abroad":   _f(best.get('חו"ל (%)')),
        "fx":       _f(best.get('מט"ח (%)')),
        "illiquid": _f(best.get("לא־סחיר (%)")),
        "sharpe":   _f(best.get("שארפ משוקלל")),
        "cost":     _f(best.get("שירות משוקלל")),
        "managers_count": (
            len(set(best.get("מנהלים", "").split("|")))
            if best.get("מנהלים") else 0
        ),
        "products_count": (
            len(best.get("weights", ())) if best.get("weights") else 0
        ),
    }

    # Changes delta
    changes: dict = {}
    for human_k, pb_k, pa_k in [
        ("מניות",    "equities", "equities"),
        ('חו"ל',     "abroad",   "abroad"),
        ('מט"ח',     "fx",       "fx"),
        ("לא-סחיר", "illiquid", "illiquid"),
    ]:
        before_v = pb.get(pb_k, _nan)
        after_v  = pa.get(pa_k, _nan)
        if not (math.isnan(before_v) or math.isnan(after_v)):
            changes[human_k] = round(after_v - before_v, 1)

    # Objectives from session state
    tgts = dict(st.session_state.get("targets", {}))
    obj = {
        "target_equities": _f(tgts.get("stocks")),
        "target_abroad":   _f(tgts.get("foreign")),
        "target_fx":       _f(tgts.get("fx")),
        "target_illiquid": _f(tgts.get("illiquid")),
        "primary_rank":    st.session_state.get("primary_rank", "דיוק"),
        "product_type":    product_type,
    }

    selected_alt_name = str(st.session_state.get("selected_alt") or best.get("חלופה", "חלופה מומלצת"))

    structured_input = {
        "portfolio_before":       pb,
        "portfolio_after":        pa,
        "client_objectives":      obj,
        "selected_solution_name": selected_alt_name,
        "changes_summary":        changes,
    }

    # ── Step 1 display: Before / After summary ────────────────────
    st.subheader("שלב 1 — השוואת תיק לפני ואחרי")
    _c1, _c2 = st.columns(2, gap="medium")
    with _c1:
        st.markdown("**📂 תיק נוכחי**")
        _rows_before = [
            ("מניות",    f"{pb['equities']:.1f}%" if not math.isnan(pb['equities'])  else "—"),
            ('חו"ל',     f"{pb['abroad']:.1f}%"   if not math.isnan(pb['abroad'])    else "—"),
            ('מט"ח',     f"{pb['fx']:.1f}%"        if not math.isnan(pb['fx'])        else "—"),
            ("לא-סחיר", f"{pb['illiquid']:.1f}%"  if not math.isnan(pb['illiquid']) else "—"),
            ("שארפ",     f"{pb['sharpe']:.2f}"     if not math.isnan(pb['sharpe'])    else "—"),
            ("עלות",     f"{pb['cost']:.2f}%"      if not math.isnan(pb['cost'])      else "—"),
        ]
        for label, val in _rows_before:
            st.markdown(f"- **{label}:** {val}")

    with _c2:
        st.markdown(f"**🎯 תיק מוצע — {selected_alt_name}**")
        _rows_after = [
            ("מניות",    f"{pa['equities']:.1f}%" if not math.isnan(pa['equities'])  else "—"),
            ('חו"ל',     f"{pa['abroad']:.1f}%"   if not math.isnan(pa['abroad'])    else "—"),
            ('מט"ח',     f"{pa['fx']:.1f}%"        if not math.isnan(pa['fx'])        else "—"),
            ("לא-סחיר", f"{pa['illiquid']:.1f}%"  if not math.isnan(pa['illiquid']) else "—"),
            ("שארפ",     f"{pa['sharpe']:.2f}"     if not math.isnan(pa['sharpe'])    else "—"),
            ("עלות",     f"{pa['cost']:.2f}%"      if not math.isnan(pa['cost'])      else "—"),
        ]
        for label, val in _rows_after:
            st.markdown(f"- **{label}:** {val}")

    if changes:
        st.markdown("**📊 שינויים עיקריים:**")
        _delta_cols = st.columns(len(changes))
        for _dc, (k, v) in zip(_delta_cols, changes.items()):
            _dc.metric(k, f"{v:+.1f}pp", delta=f"{v:+.1f}pp",
                       delta_color="normal" if v >= 0 else "inverse")

    st.divider()

    # ── Step 2: AI generation ────────────────────────────────────
    st.subheader("שלב 2 — הסבר AI (מצב: planning)")

    if "final_report_sections" not in st.session_state:
        st.session_state["final_report_sections"] = {}
    if "final_report_structured" not in st.session_state:
        st.session_state["final_report_structured"] = {}

    _ai_col, _ = st.columns([1.2, 2])
    with _ai_col:
        if st.button("🤖 הפק הסבר AI", key="btn_planning_ai", type="primary"):
            with st.spinner("מגדיר שאלת planning ושולח ל-AI..."):
                _secs, _err = run_planning_ai(structured_input)
            if _err:
                st.error(f"⚠️ שגיאת AI: {_err}")
            else:
                st.session_state["final_report_sections"]   = _secs
                st.session_state["final_report_structured"] = structured_input
                st.success("✅ הסבר AI נוצר בהצלחה")
                st.rerun()

    _secs = st.session_state.get("final_report_sections", {})

    if _secs:
        st.divider()
        # ── Step 3: User editing ──────────────────────────────────
        st.subheader("שלב 3 — עריכה ואישור")

        _tone = st.radio(
            "טון הכתיבה",
            ["מקצועי", "פשוט ונגיש", "שכנועי"],
            horizontal=True,
            key="final_report_tone",
            help="בחר טון — לאחר מכן לחץ 'הפק מחדש' לקבל טקסט מותאם",
        )

        _edited = {}
        for key in SECTION_KEYS:
            label = SECTION_LABELS_HE[key]
            default_text = _secs.get(key, "")
            _edited[key] = st.text_area(
                label,
                value=default_text,
                height=130,
                key=f"final_sec_{key}",
                help="ניתן לערוך חופשי. לחץ 'הפק מחדש' לקבל גרסה חדשה מ-AI.",
            )

        _re_col, _save_col, _ = st.columns([1, 1, 3])
        with _re_col:
            if st.button("🔄 הפק מחדש", key="btn_regen_ai", type="secondary"):
                # Adjust prompt based on tone
                _tone_instr = {
                    "מקצועי":      "כתוב בטון מקצועי ורשמי.",
                    "פשוט ונגיש": "כתוב בשפה פשוטה, נגישה ללקוח שאינו מומחה.",
                    "שכנועי":      "כתוב בטון שכנועי שמדגיש את יתרונות השינויים.",
                }.get(_tone, "")
                _mod_input = dict(structured_input)
                _mod_input["tone_instruction"] = _tone_instr
                with st.spinner("מפיק מחדש..."):
                    _new_secs, _err2 = run_planning_ai(_mod_input)
                if _err2:
                    st.error(f"⚠️ {_err2}")
                else:
                    st.session_state["final_report_sections"] = _new_secs
                    st.rerun()

        with _save_col:
            if st.button("💾 שמור עריכות", key="btn_save_edits", type="secondary"):
                st.session_state["final_report_sections"] = _edited
                st.success("עריכות נשמרו ✅")

        st.divider()

        # ── Step 4: Export ────────────────────────────────────────
        st.subheader("שלב 4 — ייצוא חבילת Notebook")

        _approved_secs = {
            k: st.session_state.get(f"final_sec_{k}", _secs.get(k, ""))
            for k in SECTION_KEYS
        }
        _pkg_str = build_notebook_package(
            structured_input, _approved_secs, product_type
        )

        _ex1, _ex2, _ = st.columns([1.2, 1.2, 4])
        with _ex1:
            st.download_button(
                "📦 הורד חבילת Notebook (JSON)",
                data=_pkg_str.encode("utf-8"),
                file_name="client_report_notebook.json",
                mime="application/json",
                key="dl_notebook_pkg",
                help="העלה קובץ זה ל-NotebookLM או ל-Colab לקבלת מצגת מלאה",
            )
        with _ex2:
            # Plain text version for review
            _plain = "\n\n".join(
                f"{SECTION_LABELS_HE[k]}\n{'-'*40}\n{_approved_secs.get(k,'')}"
                for k in SECTION_KEYS
            )
            st.download_button(
                "📄 הורד טקסט (.txt)",
                data=_plain.encode("utf-8"),
                file_name="client_report_text.txt",
                mime="text/plain",
                key="dl_text_report",
            )

        st.caption(
            "חבילת Notebook כוללת: נתוני לפני/אחרי · מטדאטה של החלופה · "
            "טקסטים מאושרים · מבנה 8 שקופיות · פרומפט מצגת"
        )
