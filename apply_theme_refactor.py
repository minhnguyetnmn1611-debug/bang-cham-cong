import os

def refactor_app_py():
    with open('app.py', 'r', encoding='utf-8') as f:
        content = f.read()

    # Replacements for global style
    content = content.replace(
        """    .vimos-fuji-overlay {
        position: fixed !important;
        top: 0 !important; left: 0 !important;
        width: 100vw !important; height: 100vh !important;
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.3) 0%, rgba(240, 249, 255, 0.45) 100%) !important;
        backdrop-filter: blur(1px) !important;
        -webkit-backdrop-filter: blur(1px) !important;
        z-index: -99998 !important;
        pointer-events: none !important;
    }""",
        """    .vimos-fuji-overlay {
        position: fixed !important;
        top: 0 !important; left: 0 !important;
        width: 100vw !important; height: 100vh !important;
        background: linear-gradient(135deg, {T['bg_app']}, {T['bg_content']}) !important;
        opacity: 0.85 !important;
        backdrop-filter: blur(2px) !important;
        -webkit-backdrop-filter: blur(2px) !important;
        z-index: -99998 !important;
        pointer-events: none !important;
    }"""
    )
    
    # Other inline styles
    content = content.replace(
        "background: {'rgba(254,252,232,0.96)' if is_sepia else 'rgba(255,255,255,0.94)'}",
        "background: {T['bg_card']}99" # Adding 99 for opacity in hex
    )
    content = content.replace(
        "border: 1.5px solid {'#D97706' if is_sepia else 'rgba(14,165,233,0.35)'}",
        "border: 1.5px solid {T['border']}"
    )
    content = content.replace(
        "background: {'linear-gradient(135deg,#B45309,#D97706)' if is_sepia else 'linear-gradient(135deg,#0EA5E9,#0284C7)'}",
        "background: {T['primary_gradient']}"
    )
    content = content.replace(
        "color: {'#92400E' if is_sepia else '#0369A1'}",
        "color: {T['text_primary']}"
    )
    content = content.replace(
        "background: {'linear-gradient(135deg, #92400E, #78350F)' if is_sepia else 'rgba(255, 255, 255, 0.94)'}",
        "background: {T['bg_card']}"
    )
    content = content.replace(
        "color: {'#FFFFFF' if is_sepia else '#0F172A'}",
        "color: {T['text_primary']}"
    )
    content = content.replace(
        "color: {'#78350F' if is_sepia else '#0F172A'}",
        "color: {T['text_primary']}"
    )
    content = content.replace(
        "text-shadow: {'none' if is_sepia else '0 1px 4px rgba(255,255,255,0.95), 0 0 10px white'}",
        "text-shadow: none"
    )
    content = content.replace(
        "color: {'#D97706' if is_sepia else '#0EA5E9'}",
        "color: {T['primary']}"
    )

    # In overview page
    content = content.replace(
        'title_color = "#92400E" if is_sepia else "#0F172A"',
        'title_color = T["text_primary"]'
    )
    content = content.replace(
        'bg_banner = "linear-gradient(135deg, rgba(254, 252, 232, 0.92) 0%, rgba(253, 246, 178, 0.88) 100%)" if is_sepia else "linear-gradient(135deg, rgba(255, 255, 255, 0.9) 0%, rgba(240, 249, 255, 0.8) 100%)"',
        'bg_banner = "transparent"' # Rely on overlay
    )
    content = content.replace(
        'banner_border = "1.5px solid rgba(217, 119, 6, 0.45)" if is_sepia else "1.5px solid rgba(14,165,233,0.3)"',
        'banner_border = f"1.5px solid {T[\'border\']}"'
    )
    content = content.replace(
        'title_color = "#92400E" if is_sepia else "#0284C7"',
        'title_color = T["primary"]'
    )
    content = content.replace(
        'subtitle_color = "#78350F" if is_sepia else "#64748B"',
        'subtitle_color = T["text_secondary"]'
    )
    content = content.replace(
        'bg_card = "rgba(254, 252, 232, 0.96)" if is_sepia else "rgba(255,255,255,0.9)"',
        'bg_card = f"{T[\'bg_card\']}E6"'
    )
    content = content.replace(
        'notif_border = "1px solid rgba(217, 119, 6, 0.45)" if is_sepia else "1px solid #E2E8F0"',
        'notif_border = f"1px solid {T[\'border\']}"'
    )
    content = content.replace(
        'title_notif_color = "#78350F" if is_sepia else "#0F172A"',
        'title_notif_color = T["text_primary"]'
    )
    content = content.replace(
        'footer_bg = "rgba(254, 252, 232, 0.98)" if is_sepia else "rgba(255, 255, 255, 0.95)"',
        'footer_bg = f"{T[\'bg_card\']}F2"'
    )
    content = content.replace(
        'footer_border = "1.5px solid rgba(217, 119, 6, 0.45)" if is_sepia else "1.5px solid rgba(14, 165, 233, 0.35)"',
        'footer_border = f"1.5px solid {T[\'border\']}"'
    )
    content = content.replace(
        "background: {'transparent' if is_sepia else 'linear-gradient(to right, #F0F9FF, transparent)'}",
        "background: {T['bg_card_hover']}"
    )
    content = content.replace(
        "color: {'#78350F' if is_sepia else '#475569'}",
        "color: {T['text_secondary']}"
    )
    content = content.replace(
        "border-left: {'2px solid #D97706' if is_sepia else '2px solid #3B82F6'}",
        "border-left: 2px solid {T['primary']}"
    )

    # In overview page summary metrics
    content = content.replace(
        'bg_banner = "linear-gradient(135deg, rgba(254, 252, 232, 0.95) 0%, rgba(253, 246, 178, 0.9) 100%)" if is_sepia else "linear-gradient(135deg, rgba(255, 255, 255, 0.95) 0%, rgba(253, 242, 248, 0.9) 100%)"',
        'bg_banner = f"{T[\'bg_card\']}E6"'
    )
    content = content.replace(
        'border_banner = "1.5px solid rgba(217, 119, 6, 0.45)" if is_sepia else "1.5px solid rgba(236,72,153,0.3)"',
        'border_banner = f"1.5px solid {T[\'border\']}"'
    )
    content = content.replace(
        'shadow_banner = "0 10px 30px rgba(120, 53, 15, 0.15)" if is_sepia else "0 10px 30px rgba(236,72,153,0.15)"',
        'shadow_banner = T["shadow"]'
    )
    content = content.replace(
        'title_color = "#92400E" if is_sepia else "#DB2777"',
        'title_color = T["primary"]'
    )
    content = content.replace(
        'text_color = "#78350F" if is_sepia else "#64748B"',
        'text_color = T["text_secondary"]'
    )
    content = content.replace(
        "background: {'rgba(254, 252, 232, 0.96)' if is_sepia else '#F8FAFC'}",
        "background: {T['bg_card_hover']}"
    )
    content = content.replace(
        "color: {'#78350F' if is_sepia else '#0F172A'}",
        "color: {T['text_primary']}"
    )
    content = content.replace(
        "color: {'#92400E' if is_sepia else '#64748B'}",
        "color: {T['text_secondary']}"
    )
    
    # Settings expander replacements
    content = content.replace(
        "background: {'rgba(254, 252, 232, 0.96)' if is_sepia else 'rgba(255,255,255,0.97)'}",
        "background: {T['bg_card']}"
    )
    content = content.replace(
        "border: 3px dashed {'rgba(217, 119, 6, 0.45)' if is_sepia else '#EC4899'}",
        "border: 3px dashed {T['border']}"
    )
    content = content.replace(
        "border-color: {'rgba(217, 119, 6, 0.65)' if is_sepia else '#0284C7'}",
        "border-color: {T['primary']}"
    )
    content = content.replace(
        "background: {'rgba(254, 240, 138, 0.96)' if is_sepia else 'rgba(255,255,255,1)'}",
        "background: {T['bg_content']}"
    )
    content = content.replace(
        "color: {'#92400E' if is_sepia else '#475569'}",
        "color: {T['text_secondary']}"
    )
    content = content.replace(
        "color: {'#D97706' if is_sepia else '#DB2777'}",
        "color: {T['primary']}"
    )
    content = content.replace(
        "background: {'linear-gradient(135deg, #F59E0B, #D97706)' if is_sepia else 'linear-gradient(135deg, #F472B6, #EC4899)'}",
        "background: {T['accent_gradient']}"
    )
    content = content.replace(
        "background: {'linear-gradient(135deg, #D97706, #B45309)' if is_sepia else 'linear-gradient(135deg, #3B82F6, #2563EB)'}",
        "background: {T['primary_gradient']}"
    )
    content = content.replace(
        "background: {'linear-gradient(135deg, #B45309 0%, #78350F 100%)' if is_sepia else 'linear-gradient(135deg, #F43F5E 0%, #E11D48 100%)'}",
        "background: {T['danger']}"
    )
    content = content.replace(
        "background: {'rgba(120, 53, 15, 0.65)' if is_sepia else 'rgba(15, 23, 42, 0.65)'}",
        "background: {T['bg_app']}D9"
    )

    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == '__main__':
    refactor_app_py()
    print("Done")
