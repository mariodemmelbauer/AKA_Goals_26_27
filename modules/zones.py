
def normalize_click(value, width=1000, height=650):
    """Return normalized 0..100 coordinates from streamlit-image-coordinates output."""
    if not value:
        return None
    x = max(0.0, min(100.0, float(value["x"]) / width * 100))
    y = max(0.0, min(100.0, float(value["y"]) / height * 100))
    return round(x, 2), round(y, 2)


def derive_zone(x, y):
    """
    Attacking direction: left -> right.
    Coordinates are normalized 0..100.
    """
    # Length bands
    if x < 33.33:
        third = "Aufbaudrittel"
    elif x < 66.67:
        third = "Mitteldrittel"
    else:
        third = "Angriffsdrittel"

    # Vertical lanes
    if y < 18:
        lane = "linker Flügel"
    elif y < 38:
        lane = "linker Halbraum"
    elif y < 62:
        lane = "Zentrum"
    elif y < 82:
        lane = "rechter Halbraum"
    else:
        lane = "rechter Flügel"

    # More specific final-third / box labels
    if x >= 83.5 and 23 <= y <= 77:
        if y < 42:
            return "Box links"
        elif y <= 58:
            return "Box zentral"
        else:
            return "Box rechts"

    if x >= 66.67 and 38 <= y <= 62:
        return "Zone 14 / Zentrum vor Box"

    return f"{third} – {lane}"
