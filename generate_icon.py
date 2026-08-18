import os
from PIL import Image, ImageDraw, ImageFont

def draw_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    pad = int(size * 0.08)
    
    # Outer rounded rectangle (Background badge)
    badge_bg = (30, 41, 59) # Dark slate header tone #1e293b
    corner_radius = int(size * 0.2)
    draw.rounded_rectangle([pad, pad, size - pad, size - pad], radius=corner_radius, fill=badge_bg)
    
    # Document shape (White page with folded corner)
    doc_left = int(size * 0.24)
    doc_top = int(size * 0.2)
    doc_right = int(size * 0.76)
    doc_bottom = int(size * 0.8)
    fold_size = int(size * 0.16)
    
    # Page points (top-left -> top-right-before-fold -> fold-bottom -> right-bottom -> left-bottom)
    page_points = [
        (doc_left, doc_top),
        (doc_right - fold_size, doc_top),
        (doc_right, doc_top + fold_size),
        (doc_right, doc_bottom),
        (doc_left, doc_bottom)
    ]
    draw.polygon(page_points, fill=(248, 250, 252)) # Crisp white #f8fafc
    
    # Fold triangle
    fold_points = [
        (doc_right - fold_size, doc_top),
        (doc_right, doc_top + fold_size),
        (doc_right - fold_size, doc_top + fold_size)
    ]
    draw.polygon(fold_points, fill=(203, 213, 225)) # Light slate fold shadow #cbd5e1
    
    # Lines inside document representing correspondence text
    line_x1 = int(size * 0.32)
    line_x2 = int(size * 0.68)
    line_color = (99, 102, 241) # Modern indigo accent #6366f1
    line_width = max(1, int(size * 0.04))
    
    # Line 1
    y1 = int(size * 0.42)
    draw.line([(line_x1, y1), (int(size * 0.6), y1)], fill=line_color, width=line_width)
    
    # Line 2
    y2 = int(size * 0.52)
    draw.line([(line_x1, y2), (line_x2, y2)], fill=line_color, width=line_width)
    
    # Line 3
    y3 = int(size * 0.62)
    draw.line([(line_x1, y3), (int(size * 0.55), y3)], fill=line_color, width=line_width)
    
    # Stamp / Seal circle at bottom right corner (Accent color #3b82f6)
    stamp_cx = int(size * 0.66)
    stamp_cy = int(size * 0.68)
    stamp_r = int(size * 0.14)
    draw.ellipse(
        [stamp_cx - stamp_r, stamp_cy - stamp_r, stamp_cx + stamp_r, stamp_cy + stamp_r],
        fill=(14, 165, 233), # Vibrant cyan/blue #0ea5e9
        outline=(255, 255, 255),
        width=max(1, int(size * 0.025))
    )
    
    # Checkmark inside stamp circle
    chk_p1 = (stamp_cx - int(stamp_r * 0.4), stamp_cy)
    chk_p2 = (stamp_cx - int(stamp_r * 0.1), stamp_cy + int(stamp_r * 0.35))
    chk_p3 = (stamp_cx + int(stamp_r * 0.45), stamp_cy - int(stamp_r * 0.35))
    draw.line([chk_p1, chk_p2, chk_p3], fill=(255, 255, 255), width=max(1, int(size * 0.04)))
    
    return img

def main():
    sizes = [256, 128, 64, 48, 32, 16]
    images = [draw_icon(s) for s in sizes]
    
    out_ico = "app_icon.ico"
    images[0].save(out_ico, format="ICO", sizes=[(s, s) for s in sizes])
    print(f"Icono generado exitosamente en {out_ico}")

if __name__ == "__main__":
    main()
