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
            
            "primary": "#EC4899",           # Xanh Blue sáng hơn cho nền tối
            "primary_hover": "#F472B6",
            "primary_gradient": "linear-gradient(135deg, #831843, #EC4899)",
            
            "accent": "#F59E0B",            # Cam hổ phách
            "accent_hover": "#FBBF24",
            "accent_gradient": "linear-gradient(135deg, #D97706, #F59E0B)",
            
            "border": "#334155",
            "border_light": "#1E293B",
            
            "shadow": "0 8px 24px rgba(0, 0, 0, 0.4)",
            "shadow_glow": "0 10px 25px -5px rgba(236, 72, 153, 0.35)", # Glow xanh
            
            "success": "#10B981",
            "warning": "#F59E0B",
            "danger": "#EF4444"
        }
    
    elif mode == 'sepia':
        return {
            "name": "sepia",
            "bg_app": "#F0E6D2",            # Nền dịu mắt hơn, giống giấy cũ
            "bg_content": "#EBE0CA",        # Nền nội dung 
            "bg_card": "#E6D9C0",           # Nền thẻ, không bị trắng chói
            "bg_card_hover": "#DFCFB3",     # Hover tối hơn một xíu
            
            "text_primary": "#4A3B32",      # Nâu sẫm
            "text_secondary": "#705A4D",    # Nâu vừa
            "text_tertiary": "#968175",     # Nâu nhạt
            
            "primary": "#B47B5A",           # Nâu đất nung (terracotta) dịu
            "primary_hover": "#9C6647",
            "primary_gradient": "linear-gradient(135deg, #9C6647, #B47B5A)",
            
            "accent": "#C48B5D",            # Cam đất
            "accent_hover": "#A67246",
            "accent_gradient": "linear-gradient(135deg, #A67246, #C48B5D)",
            
            "border": "#E8DCC8",            # Viền beige đậm hơn nền một chút
            "border_light": "#F0E6D2",
            
            "shadow": "0 4px 15px rgba(92, 71, 56, 0.08)",
            "shadow_glow": "0 10px 25px -5px rgba(180, 123, 90, 0.15)",
            
            "success": "#5C8C6B",           # Xanh lá trầm
            "warning": "#C48B5D",           # Cam đất trầm
            "danger": "#C45B5B"             # Đỏ trầm
        }
    
    else: # Default: light
        return {
            "name": "light",
            "bg_app": "#F5F6F8",            # Trắng xám rất nhạt đồng bộ tuyệt đối với sidebar
            "bg_content": "#FFFFFF",
            "bg_card": "#FFFFFF",
            "bg_card_hover": "#F8FAFC",
            
            "text_primary": "#0F172A",      # Đen slate
            "text_secondary": "#334155",    # Xám đậm
            "text_tertiary": "#64748B",     # Xám nhạt
            
            "primary": "#DB2777",           # Royal Blue
            "primary_hover": "#EC4899",
            "primary_gradient": "linear-gradient(135deg, #DB2777, #F472B6)",
            
            "accent": "#F59E0B",            # Amber
            "accent_hover": "#D97706",
            "accent_gradient": "linear-gradient(135deg, #D97706, #F59E0B)",
            
            "border": "#E2E8F0",
            "border_light": "#F1F5F9",
            
            "shadow": "0 4px 12px rgba(0, 0, 0, 0.08), 0 2px 4px rgba(0, 0, 0, 0.04)",
            "shadow_glow": "0 10px 25px -5px rgba(219, 39, 119, 0.15)", # Glow xanh nhạt
            
            "success": "#10B981",
            "warning": "#F59E0B",
            "danger": "#EF4444"
        }
