# theme.py
# Chứa cấu hình Theme (Màu sắc, Typography) cho ứng dụng VIET.MOS

def get_theme(mode: str) -> dict:
    """
    Trả về dictionary cấu hình giao diện theo 3 chế độ: 'light', 'dark', 'sepia'
    """
    if mode == 'dark':
        return {
            "name": "dark",
            "bg_app": "#0B1120",            # Nền chính của app
            "bg_content": "#0F172A",        # Nền vùng nội dung
            "bg_card": "#1E293B",           # Nền các thẻ (card)
            "bg_card_hover": "#334155",
            
            "text_primary": "#F1F5F9",      # Chữ tiêu đề, nổi bật
            "text_secondary": "#CBD5E1",    # Chữ đoạn văn thường
            "text_tertiary": "#94A3B8",     # Chữ phụ, ghi chú
            
            "primary": "#3B82F6",           # Xanh Blue sáng hơn cho nền tối
            "primary_hover": "#60A5FA",
            "primary_gradient": "linear-gradient(135deg, #1E3A8A, #3B82F6)",
            
            "accent": "#F59E0B",            # Cam hổ phách
            "accent_hover": "#FBBF24",
            "accent_gradient": "linear-gradient(135deg, #D97706, #F59E0B)",
            
            "border": "#334155",
            "border_light": "#1E293B",
            
            "shadow": "0 8px 24px rgba(0, 0, 0, 0.4)",
            "shadow_glow": "0 10px 25px -5px rgba(59, 130, 246, 0.35)", # Glow xanh
            
            "success": "#10B981",
            "warning": "#F59E0B",
            "danger": "#EF4444"
        }
    
    elif mode == 'sepia':
        return {
            "name": "sepia",
            "bg_app": "#FEFCE8",            # Vàng nhạt ấm
            "bg_content": "#FFFBEB",
            "bg_card": "#FEF3C7",
            "bg_card_hover": "#FDE68A",
            
            "text_primary": "#451A03",      # Nâu rất đậm
            "text_secondary": "#78350F",    # Nâu
            "text_tertiary": "#92400E",
            
            "primary": "#D97706",
            "primary_hover": "#B45309",
            "primary_gradient": "linear-gradient(135deg, #B45309, #D97706)",
            
            "accent": "#B45309",
            "accent_hover": "#92400E",
            "accent_gradient": "linear-gradient(135deg, #92400E, #B45309)",
            
            "border": "#FCD34D",
            "border_light": "#FDE68A",
            
            "shadow": "0 4px 15px rgba(180, 83, 9, 0.12)",
            "shadow_glow": "0 10px 25px -5px rgba(180, 83, 9, 0.25)",
            
            "success": "#059669",
            "warning": "#D97706",
            "danger": "#DC2626"
        }
    
    else: # Default: light
        return {
            "name": "light",
            "bg_app": "#F8FAFC",            # Xám nhạt ngọc trai
            "bg_content": "#FFFFFF",
            "bg_card": "#FFFFFF",
            "bg_card_hover": "#F1F5F9",
            
            "text_primary": "#0F172A",      # Đen slate
            "text_secondary": "#334155",    # Xám đậm
            "text_tertiary": "#64748B",     # Xám nhạt
            
            "primary": "#1E3A8A",           # Navy Blue
            "primary_hover": "#1E40AF",
            "primary_gradient": "linear-gradient(135deg, #1E3A8A, #2563EB)",
            
            "accent": "#F59E0B",            # Amber
            "accent_hover": "#D97706",
            "accent_gradient": "linear-gradient(135deg, #D97706, #F59E0B)",
            
            "border": "#E2E8F0",
            "border_light": "#F8FAFC",
            
            "shadow": "0 4px 6px -1px rgba(0, 0, 0, 0.05)",
            "shadow_glow": "0 10px 25px -5px rgba(37, 99, 235, 0.15)", # Glow xanh nhạt
            
            "success": "#10B981",
            "warning": "#F59E0B",
            "danger": "#EF4444"
        }
