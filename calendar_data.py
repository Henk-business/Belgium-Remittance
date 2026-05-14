# AR Calendar — Belgium AR team monthly schedule
# Updated from user's Excel calendar screenshots (November 2026 layout)
# Days are day-of-month, repeating every month

CALENDAR = {
    # Friday 1st
    1: [
        {"type": "Overview", "account": "VEPECA",        "format": "DD",     "note": "5th cutoff"},
        {"type": "Overview", "account": "BELBEV",        "format": "DD",     "note": "5th cutoff"},
    ],

    # Tuesday 4th (Mon 4.11 in image = DD VEPECA)
    # Wednesday 5th = DD BELBEV, Overview Manual PRIK & TIK
    # Friday 8th = Overview DD Vanuxeem (15th), Overview North and South beverages DD (15th),
    #              Overview Manual NEGOBOISSONS, Overview Manual Delsart,
    #              Overview DD VVD (10th), UAC PRIK & TIK

    4: [
        {"type": "DD",       "account": "VEPECA",        "format": "",       "note": ""},
    ],
    5: [
        {"type": "DD",       "account": "BELBEV",        "format": "",       "note": ""},
        {"type": "Overview", "account": "PRIK & TIK",    "format": "Manual", "note": ""},
    ],
    8: [
        {"type": "Overview", "account": "VANUXEEM",      "format": "DD",     "note": "15th cutoff"},
        {"type": "Overview", "account": "NORTH & SOUTH BEVERAGES", "format": "DD", "note": "15th cutoff"},
        {"type": "Overview", "account": "NEGOBOISSONS",  "format": "Manual", "note": "Special format see below"},
        {"type": "Overview", "account": "DELSART",       "format": "Manual", "note": ""},
        {"type": "Overview", "account": "VVD",           "format": "DD",     "note": "10th cutoff"},
        {"type": "UAC",      "account": "PRIK & TIK",    "format": "",       "note": ""},
    ],

    # Monday 11th = Overview DD VEPECA (20th), Overview DD BELBEV (20th),
    #               Overview Manual PRIK & TIK, DD VVD
    # Friday 15th = DD NASBB, DD VANUXEEM
    11: [
        {"type": "Overview", "account": "VEPECA",        "format": "DD",     "note": "20th cutoff"},
        {"type": "Overview", "account": "BELBEV",        "format": "DD",     "note": "20th cutoff"},
        {"type": "Overview", "account": "PRIK & TIK",    "format": "Manual", "note": ""},
        {"type": "DD",       "account": "VVD",           "format": "",       "note": ""},
    ],
    15: [
        {"type": "DD",       "account": "NASBB",         "format": "",       "note": ""},
        {"type": "DD",       "account": "VANUXEEM",      "format": "",       "note": ""},
    ],

    # Monday 18th = Overview DD SABIKO (28th), Overview Manual PRIK & TIK, Meeting WHS Dutch W1/W2
    # Tuesday 19th = Overview DD VVD (25th), Overview Gallee
    # Wednesday 20th = Overview DD Vanuxeem (30th), Overview Manual NEGOBOISSONS (Special format)
    18: [
        {"type": "Overview", "account": "SABIKO",        "format": "DD",     "note": "28th cutoff"},
        {"type": "Overview", "account": "PRIK & TIK",    "format": "Manual", "note": ""},
        {"type": "Meeting",  "account": "WHS Dutch",     "format": "",       "note": "W1/W2"},
    ],
    19: [
        {"type": "Overview", "account": "VVD",           "format": "DD",     "note": "25th cutoff"},
        {"type": "Overview", "account": "GALLEE",        "format": "",       "note": ""},
    ],
    20: [
        {"type": "Overview", "account": "VANUXEEM",      "format": "DD",     "note": "30th cutoff"},
        {"type": "Overview", "account": "NEGOBOISSONS",  "format": "Manual", "note": "Special format see below"},
    ],

    # Monday 25th = Overview Manual PRIK & TIK, Overview BBB (30th), Meeting WHS Dutch W1/W2, DD VVD
    # Thursday 28th = DD SABIKO
    # Friday 29th = DD VANUXEEM, DD BBB
    25: [
        {"type": "Overview", "account": "PRIK & TIK",   "format": "Manual", "note": ""},
        {"type": "Overview", "account": "BBB",           "format": "",       "note": "30th cutoff"},
        {"type": "Meeting",  "account": "WHS Dutch",     "format": "",       "note": "W1/W2"},
        {"type": "DD",       "account": "VVD",           "format": "",       "note": ""},
    ],
    28: [
        {"type": "DD",       "account": "SABIKO",        "format": "",       "note": ""},
    ],
    29: [
        {"type": "DD",       "account": "VANUXEEM",      "format": "",       "note": ""},
        {"type": "DD",       "account": "BBB",           "format": "",       "note": ""},
    ],
}

TYPE_COLORS = {
    "DD":       {"bg": "#FFC72C", "fg": "#0A0A0A", "label": "Direct Debit"},
    "Overview": {"bg": "#0A0A0A", "fg": "#FFC72C", "label": "Overview"},
    "UAC":      {"bg": "#C41230", "fg": "#FFFFFF",  "label": "UAC"},
    "Meeting":  {"bg": "#2E75B6", "fg": "#FFFFFF",  "label": "Meeting"},
}
