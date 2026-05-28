import os
import re
import json
import time
from dotenv import load_dotenv
from openai import OpenAI

# Load API key from .env
load_dotenv()
api_key = os.getenv("OPENROUTER_API_KEY")
if not api_key:
    raise ValueError("OPENROUTER_API_KEY not found in .env file")

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=api_key,
)

AI_MODEL = "google/gemini-2.5-flash-lite"

# -----------------------------------------------------------------------
# System prompt — same one used in Sheet 5
# -----------------------------------------------------------------------
SYSTEM_PROMPT = (
    "You are a GST reconciliation assistant for Indian businesses.\n\n"
    "You receive two supplier records — one from internal books, one from GSTR-2B.\n"
    "Your ONLY job is to decide if the two supplier NAMES refer to the same real-world company.\n\n"

    "WHEN TO MATCH (same company):\n"
    "- Known abbreviations or acronyms: L&T = Larsen and Toubro, HUL = Hindustan Unilever\n"
    "- Only formatting differs: M/S prefix, 'Limited'/'Ltd' suffix, '&' vs 'and', "
    "hyphen vs space, extra space, singular vs plural (Hardware vs Hardwares)\n"
    "- Clear typo in one source where the rest of the name is identical\n"
    "- Transliteration variation of the same name: Sri vs Shri, trailing vowel\n"
    "- Bank entries: internal books often record a bank as 'KBL OD 0777000100021201' "
    "(abbreviation + account type OD/CC/CA/SA + account number) while GSTR-2B has the "
    "full legal name like 'KARNATAKA BANK LIMITED'. "
    "Known bank abbreviations: SBI, KBL, HDFC, ICICI, AXIS, BOI, PNB, BOB, CANARA, "
    "KOTAK, YES, IDFC, IOB, UCO, UBI. "
    "If one name starts with a bank abbreviation followed by account details (digits, OD, CC, CA, SA) "
    "and the other is the full bank name matching that abbreviation — it is the SAME entity.\n\n"

    "WHEN NOT TO MATCH (different companies):\n"
    "- The business-type word differs: TYRES vs MOTORS, ELECTRICALS vs PLYWOOD, "
    "STEEL vs CEMENT — these identify what the business does and must agree\n"
    "- Short generic initials (NK, SR, RK) with different suffixes "
    "(Enterprises vs Industries vs Traders) — initials alone are not enough\n"
    "- Personal names with a one-letter difference and no other confirmation — "
    "SONA and SOMA are different people\n"
    "- A common first word (Akshaya, Royal, Prakash, Shri) followed by different "
    "entity words — these are likely unrelated businesses sharing a popular name\n\n"

    "DEFAULT TO NO-MATCH when uncertain. A missed match is safer than a wrong match.\n"
    "IGNORE invoice numbers and money fields completely.\n\n"

    "Respond ONLY with valid JSON, no extra text, no markdown:\n"
    '{"likely_match": true or false, "confidence": 0.0 to 1.0, "reason": "one short sentence"}'
)

# -----------------------------------------------------------------------
# Test cases — (books_supplier, 2b_supplier, expected_match, description)
# -----------------------------------------------------------------------
TEST_CASES = [
    # ---- Obvious matches ----
    (
        "L&T",
        "LARSEN AND TOUBRO LIMITED",
        True,
        "Classic abbreviation — L&T is universally known as Larsen and Toubro"
    ),
    (
        "AKSHAYA GROUPS",
        "AKSHAYA GROUP",
        True,
        "Singular vs plural — same company"
    ),
    (
        "SHRI GANESH TRADERS",
        "M/S SHRI GANESH TRADERS",
        True,
        "M/S prefix added in one source — same company"
    ),
    (
        "HITECH ENGINEERING SERVICES",
        "HI-TECH ENGINEERING SERVICES",
        True,
        "Hyphen difference in name"
    ),
    (
        "SEPL STEEL AND CEMENT",
        "SEPL STEEL & CEMENT",
        True,
        "Ampersand vs 'and'"
    ),
    (
        "VRL LOGISTICS",
        "VRL LOGISTICS LIMITED",
        True,
        "'Limited' suffix missing in books"
    ),
    (
        "MAHESH HARDWARES",
        "MAHESH HARDWARE",
        True,
        "Trailing S — singular vs plural"
    ),
    (
        "SRI VEERABHADRESHWAR",
        "SHRI VEERABHADRESHWARA",
        True,
        "Transliteration spelling variation — Sri vs Shri, trailing A"
    ),
    (
        "SAPANAPLYWOOD AND HARDWARE",
        "SAPANA PLYWOOD AND HARDWARE",
        True,
        "Space missing between words"
    ),
    (
        "KBL OD 0777000100021201",
        "KARNATAKA BANK LIMITED",
        True,
        "Bank account number used as name in books vs actual bank name in 2B"
    ),

    # ---- Clear non-matches ----
    (
        "ROYAL ENTERPRISES",
        "ROYAL TRADERS",
        False,
        "Similar first word but different second word — likely two different companies"
    ),
    (
        "BIJAPUR TYRES",
        "BIJAPUR MOTORS",
        False,
        "Same city, different business — tyres vs motors"
    ),
    (
        "NK ENTERPRISES",
        "NK INDUSTRIES",
        False,
        "Same initials, different suffix — could be different entities"
    ),
    (
        "PRAKASH ELECTRICALS",
        "PRAKASH PLYWOOD",
        False,
        "Same first name, completely different business type"
    ),
    (
        "SONA ROOPA",
        "SOMA ROOPA",
        False,
        "One-letter difference — N vs M — likely a different person"
    ),

    # ---- Borderline / tricky ----
    (
        "AKSHAYA GROUPS",
        "AKSHAYA ENTERPRISES",
        False,
        "Same first word, different entity type — could be unrelated businesses"
    ),
    (
        "SHRI SHANTESHWAR ENGINEERING",
        "SHRI SHANTESHWAR ENGINERING COMPNAY",
        True,
        "Typo in 2B — 'ENGINERING' and 'COMPNAY' are misspellings of the same company"
    ),
    (
        "KOPPAL STONE CRUSHER",
        "KOPPAL STONE CRUSHERS",
        True,
        "Trailing S — same company"
    ),
]


# -----------------------------------------------------------------------
# Run the test
# -----------------------------------------------------------------------

def ask(books_supplier, gstr2b_supplier):
    """Send one pair to the model, return parsed result."""
    payload = json.dumps({
        "books_supplier":  books_supplier,
        "2b_supplier":     gstr2b_supplier,
        "books_invoice":   "IGNORED",
        "2b_invoice":      "IGNORED",
        "taxable_diff": 0,
        "cgst_diff": 0,
        "sgst_diff": 0,
        "igst_diff": 0,
    }, ensure_ascii=False)

    response = client.chat.completions.create(
        model=AI_MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": payload},
        ],
        temperature=0,
    )
    raw = response.choices[0].message.content.strip()
    clean = re.sub(r"^```[a-z]*\n?", "", raw).rstrip("` \n")
    return json.loads(clean)


print(f"Model : {AI_MODEL}")
print(f"Cases : {len(TEST_CASES)}\n")
print(f"{'#':<3} {'BOOKS SUPPLIER':<35} {'2B SUPPLIER':<40} {'EXP':>5} {'GOT':>5} {'CONF':>6}  {'PASS':>5}  REASON")
print("-" * 150)

passed = 0
failed = 0
results = []

for i, (books_sup, gstr2b_sup, expected, description) in enumerate(TEST_CASES, 1):
    try:
        result = ask(books_sup, gstr2b_sup)
        got        = result["likely_match"]
        confidence = result["confidence"]
        reason     = result["reason"]
        ok         = (got == expected)

        if ok:
            passed += 1
            status = "PASS"
        else:
            failed += 1
            status = "FAIL ←"

        print(
            f"{i:<3} {books_sup:<35} {gstr2b_sup:<40} "
            f"{'T' if expected else 'F':>5} "
            f"{'T' if got else 'F':>5} "
            f"{confidence:>6.2f}  {status:>6}  {reason}"
        )
        results.append({
            "books": books_sup, "gstr2b": gstr2b_sup,
            "expected": expected, "got": got,
            "confidence": confidence, "reason": reason,
            "pass": ok, "description": description,
        })

    except Exception as e:
        failed += 1
        print(f"{i:<3} {books_sup:<35} {gstr2b_sup:<40} ERROR: {e}")

    time.sleep(0.3)  # rate limit

# -----------------------------------------------------------------------
# Summary
# -----------------------------------------------------------------------
print("\n" + "=" * 60)
print(f"Results: {passed} passed, {failed} failed out of {len(TEST_CASES)}")
print(f"Accuracy: {passed / len(TEST_CASES) * 100:.1f}%")

if failed > 0:
    print("\nFailed cases:")
    for r in results:
        if not r["pass"]:
            exp = "match" if r["expected"] else "no-match"
            got = "match" if r["got"] else "no-match"
            print(f"  [{exp} expected, got {got}] {r['books']} ↔ {r['gstr2b']}")
            print(f"    → {r['description']}")
