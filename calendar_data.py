# AR Calendar — parsed from user's monthly schedule
# Structure: day_of_month -> list of tasks
# Tasks have: type (DD/Overview/Retour/UAC/Meeting), account, notes

CALENDAR = {
    1: [
        {"type": "Overview", "account": "VEPECA",        "format": "DD",     "note": "5th cutoff"},
        {"type": "Retour",   "account": "VEPECA",        "format": "",       "note": ""},
        {"type": "Overview", "account": "BELBEV",        "format": "DD",     "note": "5th cutoff"},
        {"type": "Retour",   "account": "BELBEV",        "format": "",       "note": ""},
    ],
    4: [
        {"type": "DD",       "account": "VEPECA",        "format": "",       "note": ""},
    ],
    5: [
        {"type": "DD",       "account": "BELBEV",        "format": "",       "note": ""},
        {"type": "Overview", "account": "PRIK & TIK",    "format": "Manual", "note": ""},
        {"type": "Overview", "account": "VANUXEEM",      "format": "DD",     "note": "15th cutoff"},
    ],
    6: [
        {"type": "Retour",   "account": "VANUXEEM",      "format": "",       "note": ""},
        {"type": "Overview", "account": "NORTH & SOUTH BEVERAGES", "format": "DD", "note": "15th cutoff"},
        {"type": "Retour",   "account": "NORTH & SOUTH BEVERAGES","format": "", "note": ""},
        {"type": "Overview", "account": "NEGOBOISSONS",  "format": "Manual", "note": "Special format"},
        {"type": "Overview", "account": "DELSART",       "format": "Manual", "note": ""},
        {"type": "Overview", "account": "VVD",           "format": "DD",     "note": "10th cutoff"},
        {"type": "UAC",      "account": "PRIK & TIK",    "format": "",       "note": ""},
    ],
    11: [
        {"type": "DD",       "account": "NASBB",         "format": "",       "note": ""},
        {"type": "DD",       "account": "VANUXEEM",      "format": "",       "note": ""},
        {"type": "Overview", "account": "VEPECA",        "format": "DD",     "note": "20th cutoff"},
        {"type": "Retour",   "account": "VEPECA",        "format": "",       "note": ""},
        {"type": "Overview", "account": "BELBEV",        "format": "DD",     "note": "20th cutoff"},
        {"type": "Retour",   "account": "BELBEV",        "format": "",       "note": ""},
        {"type": "Overview", "account": "PRIK & TIK",    "format": "Manual", "note": ""},
        {"type": "DD",       "account": "VVD",           "format": "",       "note": ""},
    ],
    18: [
        {"type": "Overview", "account": "SABIKO",        "format": "DD",     "note": "28th cutoff"},
        {"type": "Retour",   "account": "SABIKO",        "format": "",       "note": ""},
        {"type": "Overview", "account": "VVD",           "format": "DD",     "note": "25th cutoff"},
        {"type": "Overview", "account": "VANUXEEM",      "format": "DD",     "note": "30th cutoff"},
        {"type": "Retour",   "account": "VANUXEEM",      "format": "",       "note": ""},
        {"type": "Overview", "account": "GALLEE",        "format": "",       "note": ""},
        {"type": "Overview", "account": "PRIK & TIK",    "format": "Manual", "note": ""},
        {"type": "Overview", "account": "NEGOBOISSONS",  "format": "Manual", "note": "Special format"},
        {"type": "Meeting",  "account": "WHS Dutch",     "format": "",       "note": "W1/W2"},
    ],
    25: [
        {"type": "Overview", "account": "PRIK & TIK",   "format": "Manual", "note": ""},
        {"type": "DD",       "account": "SABIKO",        "format": "",       "note": ""},
        {"type": "DD",       "account": "VANUXEEM",      "format": "",       "note": ""},
        {"type": "DD",       "account": "BBB",           "format": "",       "note": ""},
        {"type": "Overview", "account": "BBB",           "format": "",       "note": "30th cutoff"},
        {"type": "Meeting",  "account": "WHS Dutch",     "format": "",       "note": "W1/W2"},
        {"type": "DD",       "account": "VVD",           "format": "",       "note": ""},
    ],
}

# Colour coding per task type
TYPE_COLORS = {
    "DD":       {"bg": "#FFC72C", "fg": "#0A0A0A", "label": "Direct Debit"},
    "Overview": {"bg": "#0A0A0A", "fg": "#FFC72C", "label": "Overview"},
    "Retour":   {"bg": "#E8E3DC", "fg": "#3A3530", "label": "Retour"},
    "UAC":      {"bg": "#C41230", "fg": "#FFFFFF",  "label": "UAC"},
    "Meeting":  {"bg": "#2E75B6", "fg": "#FFFFFF",  "label": "Meeting"},
}
