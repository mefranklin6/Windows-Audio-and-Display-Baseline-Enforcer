from PIL import Image, ImageDraw

def generate_dual_monitor_icon(target_size=512):
    # Supersampling 2x for sharp anti-aliasing
    scale = 2
    size = target_size * scale
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    s = size / 512.0

    # 1. Outer Background Squircle Tile
    padding = int(16 * s)
    radius = int(96 * s)
    draw.rounded_rectangle(
        [padding, padding, size - padding, size - padding],
        radius=radius, fill="#0f172a"
    )
    draw.rounded_rectangle(
        [padding, padding, size - padding, size - padding],
        radius=radius, outline="#0284c7", width=int(4 * s)
    )

    baseline_y = int(220 * s)

    # 2. Left Monitor (Portrait Mode)
    p_left, p_top = int(80 * s), int(100 * s)
    p_right, p_bottom = int(210 * s), int(340 * s)

    # Stand (Portrait)
    draw.polygon([
        (int(115*s), int(380*s)), (int(175*s), int(380*s)),
        (int(165*s), int(340*s)), (int(125*s), int(340*s))
    ], fill="#334155")
    draw.rectangle([int(135*s), int(340*s), int(155*s), int(365*s)], fill="#1e293b")

    # Bezel & Screen (Portrait)
    draw.rounded_rectangle(
        [p_left, p_top, p_right, p_bottom],
        radius=int(14 * s), fill="#1e293b", outline="#475569", width=int(4 * s)
    )
    sp = int(10 * s)
    draw.rounded_rectangle(
        [p_left + sp, p_top + sp, p_right - sp, p_bottom - sp],
        radius=int(8 * s), fill="#020617"
    )

    # 3. Right Monitor (Landscape Mode)
    l_left, l_top = int(225 * s), int(130 * s)
    l_right, l_bottom = int(435 * s), int(320 * s)

    # Stand (Landscape)
    draw.polygon([
        (int(290*s), int(360*s)), (int(370*s), int(360*s)),
        (int(360*s), int(320*s)), (int(300*s), int(320*s))
    ], fill="#334155")
    draw.rectangle([int(315*s), int(320*s), int(345*s), int(345*s)], fill="#1e293b")

    # Bezel & Screen (Landscape)
    draw.rounded_rectangle(
        [l_left, l_top, l_right, l_bottom],
        radius=int(14 * s), fill="#1e293b", outline="#475569", width=int(4 * s)
    )
    draw.rounded_rectangle(
        [l_left + sp, l_top + sp, l_right - sp, l_bottom - sp],
        radius=int(8 * s), fill="#020617"
    )

    # 4. Audio Equalizer Waveform Bars (Spanning across both displays)
    # Format: (center_x, height, color)
    bars = [
        # Portrait Display Bars
        (105, 30, "#06b6d4"),
        (125, 80, "#0ea5e9"),
        (145, 140, "#3b82f6"),
        (165, 90, "#0ea5e9"),
        (185, 50, "#06b6d4"),

        # Landscape Display Bars
        (250, 60, "#06b6d4"),
        (270, 110, "#0ea5e9"),
        (290, 150, "#3b82f6"),
        (310, 80, "#0ea5e9"),
        (330, 120, "#3b82f6"),
        (350, 70, "#0ea5e9"),
        (370, 30, "#06b6d4"),
        (390, 16, "#06b6d4")
    ]

    bar_width = int(12 * s)
    for cx, bh, color in bars:
        x1 = int((cx - 6) * s)
        h = int(bh * s)
        y1 = baseline_y - (h // 2)
        draw.rounded_rectangle([x1, y1, x1 + bar_width, y1 + h], radius=int(6 * s), fill=color)

    # 5. Baseline Dotted Reference Line
    for x in range(int(90 * s), int(425 * s), int(12 * s)):
        in_portrait = (int(90 * s) <= x <= int(200 * s))
        in_landscape = (int(235 * s) <= x <= int(425 * s))
        if in_portrait or in_landscape:
            draw.line([(x, baseline_y), (x + int(6 * s), baseline_y)], fill="#38bdf8", width=int(3 * s))

    # Live Monitoring Status Dot (Landscape Display)
    draw.ellipse([int(405*s), int(148*s), int(419*s), int(162*s)], fill="#10b981")
    draw.ellipse([int(402*s), int(145*s), int(422*s), int(165*s)], fill=None, outline="#34d399", width=int(2*s))

    # 6. Enforcement Shield Badge
    shield_pts = [
        (int(360*s), int(285*s)), (int(425*s), int(315*s)),
        (int(425*s), int(370*s)), (int(360*s), int(440*s)),
        (int(295*s), int(370*s)), (int(295*s), int(315*s))
    ]
    draw.polygon(shield_pts, fill="#059669", outline="#34d399", width=int(4 * s))

    # Checkmark inside Shield
    chk = [(int(332*s), int(368*s)), (int(352*s), int(388*s)), (int(390*s), int(348*s))]
    draw.line([chk[0], chk[1]], fill="#ffffff", width=int(8 * s))
    draw.line([chk[1], chk[2]], fill="#ffffff", width=int(8 * s))

    # Downsample Lanczos for final icon output
    return img.resize((target_size, target_size), Image.Resampling.LANCZOS)

if __name__ == "__main__":
    icon = generate_dual_monitor_icon(512)
    icon.save("dual_monitor_icon.png")
    icon.save("app_icon.ico", format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("Exported dual_monitor_icon.png & app_icon.ico successfully!")