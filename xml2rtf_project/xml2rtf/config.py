STYLE_MAP = {
    "Title_document": r"\fs36\b\qc",  # 18pt, bold, centered
    "Body_text": r"\fs24\ql",         # 12pt, left aligned
}

INLINE_MAP = {
    "bold": {"true": r"\b", "false": ""},
    "italic": {"true": r"\i", "false": ""},
    "underline": {"true": r"\ul", "false": ""},
    "color": {
        "red": r"\cf1",  # Color table index
        "blue": r"\cf2",
        # ... add more as needed
    },
    # Add font, size, etc.
}

COLOR_TABLE = r"{\colortbl ;\red255\green0\blue0;\red0\green0\blue255;}"
