import os
import pandas as pd
from app import (
    clean_phone_number,
    format_template_message,
    load_blocklist,
    append_to_blocklist,
    load_sent_history,
    append_to_sent_history,
    DEFAULT_MESSAGE_TEMPLATE,
    COMPANY_NAME,
    COMPANY_WEBSITE,
    COMPANY_EMAIL,
    SENDER_PHONE,
)

def test_clean_phone():
    test_cases = [
        ("9812345670", "+919812345670", None),
        ("9834567892.0", "+919834567892", None),
        ("+91 98567 89014", "+919856789014", None),
        ("09812345670", "+919812345670", None),
        ("919812345670", "+919812345670", None),
        ("+14155552671", "+14155552671", None),
        ("12345", None, "Invalid"),
        ("", None, "Empty"),
        (None, None, "Empty"),
        (float("nan"), None, "Empty"),
    ]
    for raw, expected_phone, expected_err in test_cases:
        phone, err = clean_phone_number(raw)
        if expected_phone:
            assert phone == expected_phone, f"Failed for '{raw}': got '{phone}', expected '{expected_phone}'"
            assert err is None
        else:
            assert phone is None, f"Expected None for '{raw}', got '{phone}'"
            assert expected_err.lower() in err.lower()
    print("[PASS] test_clean_phone passed!")

def test_formatting():
    contact = {
        "Name": "Aarav Sharma",
        "Company": "Apex Retail",
        "Service": "E-Commerce Website",
        "Suggestion": "Speed up mobile checkout to increase conversion rate by 30%.",
    }
    msg = format_template_message(DEFAULT_MESSAGE_TEMPLATE, contact)
    assert "Hi Aarav Sharma," in msg
    assert "Apex Retail" in msg
    assert "E-Commerce Website" in msg
    assert "Speed up mobile checkout" in msg
    assert "https://skytecher.com" in msg
    assert "skytechersolutions@gmail.com" in msg
    assert "8960061745" in msg
    assert "Reply STOP" in msg
    print("[PASS] test_formatting passed with full Skytecher details!")

def test_env_defaults():
    assert COMPANY_NAME == "Skytecher"
    assert "skytecher.com" in COMPANY_WEBSITE
    assert COMPANY_EMAIL == "skytechersolutions@gmail.com"
    assert SENDER_PHONE == "8960061745"
    print("[PASS] test_env_defaults passed!")

def test_excel_loading():
    df = pd.read_excel("sample_leads.xlsx")
    assert len(df) == 7
    expected_cols = ["#", "Company", "Industry", "Lead Type", "Country", "City", "Phone", "Website", "Website Status", "LinkedIn (company)", "Google Rating", "Source", "Suggested Services", "Contact Score", "Priority", "Outreach Status", "Notes"]
    for col in expected_cols:
        assert col in df.columns, f"Missing column: {col}"
    print(f"[PASS] sample_leads.xlsx loaded with {len(df)} rows and all 17 columns!")

def test_17_column_mapping():
    df = pd.read_excel("sample_leads.xlsx")
    row = df.iloc[0]
    contact = {
        "Company": row["Company"],
        "Suggested Services": row["Suggested Services"],
        "Notes": row["Notes"],
        "City": row["City"],
        "Industry": row["Industry"],
    }
    msg = format_template_message(DEFAULT_MESSAGE_TEMPLATE, contact)
    assert "Apex Retail Solutions" in msg
    assert "Modern Website Development" in msg
    assert "lightning-fast mobile website" in msg
    assert "https://skytecher.com" in msg
    print("[PASS] 17-column lead format successfully mapped to template!")

if __name__ == "__main__":
    test_clean_phone()
    test_formatting()
    test_env_defaults()
    test_excel_loading()
    test_17_column_mapping()
    print("\nAll functional tests passed successfully!")
