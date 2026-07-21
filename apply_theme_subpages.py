import os

def refactor_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Add import if missing
    if 'from theme import get_theme' not in content:
        content = content.replace("import streamlit as st", "import streamlit as st\nfrom theme import get_theme")

    # Replace is_sepia with theme logic
    content = content.replace(
        "is_sepia = st.session_state.get('eye_care_sepia', False)",
        "theme_mode = st.session_state.get('theme_mode', 'light')\n    gT = get_theme(theme_mode)\n    is_sepia = (theme_mode == 'sepia')"
    )

    # In page_mos.py, replace the T dictionary
    if 'page_mos.py' in filepath:
        old_T_dict = """    T = {
        # Backgrounds
        "bg_header":     "rgba(254,252,232,0.95)"        if is_sepia else "rgba(255,255,255,0.85)",
        "bg_card":       "rgba(254,252,232,0.96)"        if is_sepia else "rgba(255,255,255,0.95)",
        "bg_section":    "#FEFCE8"                       if is_sepia else "rgba(255,255,255,0.85)",
        "bg_input":      "rgba(254,252,232,0.92)"        if is_sepia else "rgba(255,255,255,0.9)",
        "bg_tag":        "#FDE68A"                       if is_sepia else "rgba(14,165,233,0.08)",
        "bg_result":     "#FEFCE8"                       if is_sepia else "rgba(255,255,255,0.97)",
        "bg_upload":     "#FEFCE8"                       if is_sepia else "rgba(255,255,255,0.9)",
        "bg_step_active":"#FDE68A"                       if is_sepia else "rgba(14,165,233,0.08)",
        # Borders
        "border_header": "rgba(217,119,6,0.45)"          if is_sepia else "rgba(236,72,153,0.35)",
        "border_card":   "rgba(217,119,6,0.35)"          if is_sepia else "rgba(14,165,233,0.2)",
        "border_input":  "rgba(217,119,6,0.4)"           if is_sepia else "#CBD5E1",
        "border_upload": "rgba(217,119,6,0.4)"           if is_sepia else "#93C5FD",
        "border_result": "rgba(217,119,6,0.35)"          if is_sepia else "rgba(14,165,233,0.2)",
        # Text
        "text_h1":       "#78350F"                       if is_sepia else "#0F172A",
        "text_h2":       "#78350F"                       if is_sepia else "#1E293B",
        "text_body":     "#92400E"                       if is_sepia else "#334155",
        "text_muted":    "#92400E"                       if is_sepia else "#64748B",
        "text_accent":   "#D97706"                       if is_sepia else "#0EA5E9",
        "text_label":    "#78350F"                       if is_sepia else "#334155",
        # Shadows
        "shadow_header": "rgba(120,53,15,0.15)"          if is_sepia else "rgba(236,72,153,0.15)",
        "shadow_card":   "rgba(120,53,15,0.08)"          if is_sepia else "rgba(14,165,233,0.08)",
        # Accents
        "accent_btn":    "#D97706"                       if is_sepia else "#0EA5E9",
        "accent_btn_hover": "#B45309"                    if is_sepia else "#0284C7",
        "accent_border_focus": "#D97706"                 if is_sepia else "#0EA5E9",
        "pulse_color":   "#D97706"                       if is_sepia else "#0EA5E9",
    }"""
        new_T_dict = """    T = {
        "bg_header":     f"{gT['bg_card']}E6",
        "bg_card":       f"{gT['bg_card']}F2",
        "bg_section":    gT['bg_app'],
        "bg_input":      f"{gT['bg_content']}E6",
        "bg_tag":        f"{gT['primary']}1A",
        "bg_result":     gT['bg_app'],
        "bg_upload":     gT['bg_app'],
        "bg_step_active":f"{gT['primary']}1A",
        "border_header": gT['border'],
        "border_card":   gT['border'],
        "border_input":  gT['border'],
        "border_upload": gT['primary'],
        "border_result": gT['border'],
        "text_h1":       gT['text_primary'],
        "text_h2":       gT['text_primary'],
        "text_body":     gT['text_secondary'],
        "text_muted":    gT['text_tertiary'],
        "text_accent":   gT['primary'],
        "text_label":    gT['text_secondary'],
        "shadow_header": gT['shadow'],
        "shadow_card":   gT['shadow'],
        "accent_btn":    gT['primary'],
        "accent_btn_hover": gT['primary_hover'],
        "accent_border_focus": gT['primary'],
        "pulse_color":   gT['primary'],
    }"""
        content = content.replace(old_T_dict, new_T_dict)

        # Inline replace in page_mos CSS
        content = content.replace("{'rgba(254, 252, 232, 0.95)' if is_sepia else 'rgba(255, 255, 255, 0.85)'}", "{gT['bg_card']}E6")
        content = content.replace("{'rgba(217, 119, 6, 0.45)' if is_sepia else 'rgba(236, 72, 153, 0.35)'}", "{gT['border']}")
        content = content.replace("{'rgba(120, 53, 15, 0.15)' if is_sepia else 'rgba(236, 72, 153, 0.15)'}", "{gT['shadow']}")
        content = content.replace("{'rgba(217, 119, 6, 0.1)' if is_sepia else 'rgba(14, 165, 233, 0.05)'}", "{gT['primary']}1A")
        content = content.replace("{'#78350F' if is_sepia else '#1E293B'}", "{gT['text_primary']}")
        content = content.replace('{"" if is_sepia else "background: linear-gradient(135deg, #EC4899 0%, #0284C7 100%);"}', "background: {gT['primary_gradient']};")
        content = content.replace("{'#92400E' if is_sepia else '#64748B'}", "{gT['text_secondary']}")
        content = content.replace("{'rgba(217, 119, 6, 0.15)' if is_sepia else 'rgba(236, 72, 153, 0.15)'}", "{gT['shadow']}")
        content = content.replace("{'#92400E' if is_sepia else 'white'}", "{gT['text_primary']}")
        content = content.replace("{'rgba(254, 252, 232, 0.94)' if is_sepia else 'rgba(255, 255, 255, 0.94)'}", "{gT['bg_card']}F2")
        content = content.replace("{'rgba(217, 119, 6, 0.25)' if is_sepia else 'rgba(14, 165, 233, 0.2)'}", "{gT['border']}")
        content = content.replace("{'#FEFCE8' if is_sepia else 'white'}", "{gT['bg_content']}")
        content = content.replace("{'#D97706' if is_sepia else '#0284C7'}", "{gT['primary']}")
        content = content.replace("{'rgba(217, 119, 6, 0.18)' if is_sepia else 'rgba(2, 132, 199, 0.18)'}", "{gT['primary']}2E")
        content = content.replace("{'rgba(217, 119, 6, 0.35)' if is_sepia else 'rgba(2, 132, 199, 0.35)'}", "{gT['primary']}59")
        content = content.replace("{'rgba(217, 119, 6, 0)' if is_sepia else 'rgba(2, 132, 199, 0)'}", "{gT['primary']}00")
        content = content.replace("{'#FDE68A' if is_sepia else '#CBD5E1'}", "{gT['border']}")
        content = content.replace("{'#B45309' if is_sepia else '#475569'}", "{gT['text_secondary']}")
        content = content.replace("{'#F59E0B' if is_sepia else '#0EA5E9'}", "{gT['primary']}")
        content = content.replace("{'rgba(254, 252, 232, 0.95)' if is_sepia else 'rgba(255, 255, 255, 0.9)'}", "{gT['bg_card']}E6")
        content = content.replace("{'#FDE68A' if is_sepia else '#E2E8F0'}", "{gT['border']}")
        content = content.replace("{'#78350F' if is_sepia else '#334155'}", "{gT['text_secondary']}")
        content = content.replace("{'#D97706' if is_sepia else '#BFDBFE'}", "{gT['primary']}")
        content = content.replace("{'#FEF3C7' if is_sepia else '#F1F5F9'}", "{gT['bg_app']}")
        content = content.replace("{'#78350F' if is_sepia else '#0F172A'}", "{gT['text_primary']}")
        content = content.replace("{'#FEFCE8' if is_sepia else '#F8FAFC'}", "{gT['bg_app']}")
        content = content.replace("{'#D97706' if is_sepia else '#0EA5E9'}", "{gT['primary']}")
        content = content.replace("{'#B45309' if is_sepia else '#0284C7'}", "{gT['primary_hover']}")
        content = content.replace("{'rgba(217, 119, 6, 0.4)' if is_sepia else 'rgba(14, 165, 233, 0.4)'}", "{gT['shadow']}")

    # Apply to page_gamification and page_checkin too (they just need is_sepia to be gT keys)
    if 'page_gamification.py' in filepath:
        content = content.replace("{'rgba(254, 252, 232, 0.95)' if is_sepia else 'rgba(255, 255, 255, 0.95)'}", "{gT['bg_card']}")
        content = content.replace("{'#92400E' if is_sepia else '#0F172A'}", "{gT['text_primary']}")
        content = content.replace("{'#78350F' if is_sepia else '#475569'}", "{gT['text_secondary']}")
        content = content.replace("{'1.5px solid rgba(217, 119, 6, 0.45)' if is_sepia else '1px solid rgba(244, 114, 182, 0.3)'}", "1.5px solid {gT['border']}")
        content = content.replace("{'rgba(254, 240, 138, 0.96)' if is_sepia else 'rgba(240, 249, 255, 0.95)'}", "{gT['bg_app']}")
        content = content.replace("{'#D97706' if is_sepia else '#EC4899'}", "{gT['primary']}")
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == '__main__':
    for p in ['page_mos.py', 'page_gamification.py', 'page_checkin.py']:
        if os.path.exists(p):
            refactor_file(p)
    print("Done refactoring subpages")
