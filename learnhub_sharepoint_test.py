# learnhub_sharepoint_test.py
# Tests: (1) auth, (2) write a List item, (3) upload a small test file
# Requirements: pip install msal requests

import msal
import requests
import json
import os
from datetime import datetime

# ── CONFIG — same values as your HTML file ────────────────────────────────────
CONFIG = {
    "CLIENT_ID":           "4ab47c44-464a-479f-b65b-e5b3669514e1",
    "TENANT_ID":           "254ba93e-1f6f-48f3-90e6-e2766664b477",
    "SITE_ID":             "siemensenergyag.sharepoint.com,4a910703-8747-4f36-b737-1542551ae10d,08e58f6a-9127-4271-a62d-b47bed427d4c",
    "LIST_ID":             "3a644fea-8093-4bb4-8245-a8753f423433",
    "DRIVE_ID":            "b!AweRSkeHNk-3NxVCVRrhDWqP5QgnkXFCpi20e-1CfUyVj7PYHaOQTo1KbSJu7Fd1",
    "LIBRARY_ROOT_FOLDER": "GC_Training_Portal_Test/Training Materials Library",
}

# App-only auth: read the client secret from an env var, never hardcode it.
CLIENT_SECRET = os.environ["SP_CLIENT_SECRET"]

SCOPES = ["https://graph.microsoft.com/.default"]
GRAPH  = "https://graph.microsoft.com/v1.0"

# ── AUTH — client credentials flow (app-only, no user sign-in) ───────────────
def get_token():
    app = msal.ConfidentialClientApplication(
        CONFIG["CLIENT_ID"],
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{CONFIG['TENANT_ID']}",
    )
    # Try silent first (cached token)
    result = app.acquire_token_silent(SCOPES, account=None)
    if result and "access_token" in result:
        print("✓ Token acquired silently (cached)")
        return result["access_token"]

    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        raise Exception(f"Auth failed: {result.get('error_description', result)}")
    print("✓ App-only token acquired\n")
    return result["access_token"]

# ── GRAPH HELPERS ─────────────────────────────────────────────────────────────
def graph_get(token, path):
    r = requests.get(f"{GRAPH}{path}", headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    return r.json()

def graph_post(token, path, body):
    r = requests.post(
        f"{GRAPH}{path}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
    )
    r.raise_for_status()
    return r.json()

def graph_put_file(token, path, data, content_type="application/octet-stream"):
    r = requests.put(
        f"{GRAPH}{path}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": content_type},
        data=data,
    )
    r.raise_for_status()
    return r.json()

# ── TEST 0: Create a new SharePoint List and return its ID ───────────────────
def test_create_list(token, name="LearnHub Test List"):
    print("── Test 0: Create SharePoint List ──────────────────")
    body = {
        "displayName": name,
        "columns": [
            {"name": "Description",     "text": {}},
            {"name": "SubmittedBy",     "text": {}},
            {"name": "Team",            "text": {}},
            {"name": "Category",        "text": {}},
            {"name": "ContentType",     "text": {}},
            {"name": "Status",          "text": {}},
            {"name": "ReferenceNumber", "text": {}},
        ],
        "list": {"template": "genericList"},
    }
    lst = graph_post(token, f"/sites/{CONFIG['SITE_ID']}/lists", body)
    print(f"  ✓ List created — Name: {lst['displayName']}")
    print(f"  ✓ LIST_ID      : {lst['id']}")
    print(f"  ✓ List URL     : {lst['webUrl']}")
    print("  → Copy this LIST_ID into CONFIG['LIST_ID'] to reuse it.\n")
    return lst["id"]

# ── TEST 1: Verify site and list are reachable ────────────────────────────────
def test_read_access(token):
    print("── Test 1: Read access ──────────────────────────────")
    site = graph_get(token, f"/sites/{CONFIG['SITE_ID']}")
    print(f"  Site name    : {site['displayName']}")
    print(f"  Site URL     : {site['webUrl']}")

    lst = graph_get(token, f"/sites/{CONFIG['SITE_ID']}/lists/{CONFIG['LIST_ID']}")
    print(f"  List name    : {lst['displayName']}")
    print(f"  List item ct : {lst.get('list', {}).get('itemCount', '?')} existing items")

    drive = graph_get(token, f"/sites/{CONFIG['SITE_ID']}/drives/{CONFIG['DRIVE_ID']}")
    print(f"  Library name : {drive['name']}")
    print("  ✓ Read access confirmed\n")

# ── TEST 2: Write a test item to the SharePoint List ─────────────────────────
def test_list_write(token):
    print("── Test 2: Write to SharePoint List ────────────────")
    ref = f"LH-TEST-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    item = graph_post(token, f"/sites/{CONFIG['SITE_ID']}/lists/{CONFIG['LIST_ID']}/items", {
        "fields": {
            "Title":            "TEST — Python write test",
            "Description":      "Automated test item written by learnhub_sharepoint_test.py",
            "SubmittedBy":      "Python Test Script",
            "Team":             "GC UK",
            "Category":         "Technical Skills",
            "ContentType":      "Document / Guide",
            "Status":           "Pending Review",
            "ReferenceNumber":  ref,
        }
    })
    item_id = item["id"]
    print(f"  ✓ List item created — ID: {item_id}, Ref: {ref}")

    # Clean up: delete the test item
    requests.delete(
        f"{GRAPH}/sites/{CONFIG['SITE_ID']}/lists/{CONFIG['LIST_ID']}/items/{item_id}",
        headers={"Authorization": f"Bearer {token}"},
    )
    print("  ✓ Test item deleted (cleanup)\n")
    return True

# ── TEST 3: Upload a small test file to the Document Library ──────────────────
def test_file_upload(token):
    print("── Test 3: Upload test file to Document Library ────")
    folder   = CONFIG["LIBRARY_ROOT_FOLDER"] + "/Technical Skills"
    filename = f"test_upload_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt"
    content  = b"LearnHub SharePoint write test. Safe to delete."

    encoded_path = f"{folder}/{filename}".replace(" ", "%20")
    item = graph_put_file(
        token,
        f"/sites/{CONFIG['SITE_ID']}/drives/{CONFIG['DRIVE_ID']}/root:/{encoded_path}:/content",
        data=content,
        content_type="text/plain",
    )
    print(f"  ✓ File uploaded  : {item['name']}")
    print(f"  ✓ SharePoint URL : {item['webUrl']}")

    # Clean up: delete the test file
    requests.delete(
        f"{GRAPH}/sites/{CONFIG['SITE_ID']}/drives/{CONFIG['DRIVE_ID']}/items/{item['id']}",
        headers={"Authorization": f"Bearer {token}"},
    )
    print("  ✓ Test file deleted (cleanup)\n")
    return True

# ── ALSO USEFUL: discover Site ID if you don't have it yet ───────────────────
def discover_site_id(token, hostname, site_path):
    """
    Call this if you need to find your SITE_ID.
    hostname  = "siemensenergyag.sharepoint.com"
    site_path = "/teams/M_Digitial_Grid_Documents"
    """
    result = graph_get(token, f"/sites/{hostname}:{site_path}")
    print(f"\nSite display name : {result['displayName']}")
    print(f"SITE_ID to use    : {result['id']}\n")
    return result["id"]

def create_training_list(token):
    print("── Creating Training Repository list ───────────────")

    # 1. Create the list
    lst = graph_post(token, f"/sites/{CONFIG['SITE_ID']}/lists", {
        "displayName": "Training Repository",
        "columns": [],
        "list": { "template": "genericList" }
    })
    list_id = lst["id"]
    print(f"  ✓ List created — ID: {list_id}")

    # 2. Add all columns
    columns = [
        {"name": "Description",      "text": {}},
        {"name": "SubmittedBy",       "text": {}},
        {"name": "SubmitterEmail",    "text": {}},
        {"name": "Team",              "text": {}},
        {"name": "DateCreated",       "dateTime": {"format": "dateOnly"}},
        {"name": "Version",           "text": {}},
        {"name": "MaterialType",      "choice": {"choices": ["E-learning Module","Video","Document / Guide","Presentation","Job Aid / Quick Reference","Webinar Recording","External Course Link"]}},
        {"name": "Category",          "choice": {"choices": ["Compliance & Legal","Leadership & Management","Technical Skills","Soft Skills","Health & Safety","Onboarding","Data & Analytics","Diversity & Inclusion"]}},
        {"name": "TargetAudience",    "text": {"allowMultipleLines": True}},
        {"name": "DifficultyLevel",   "choice": {"choices": ["Introductory","Intermediate","Advanced","All Levels"]}},
        {"name": "Language",          "choice": {"choices": ["English","French","German","Spanish","Mandarin","Other"]}},
        {"name": "ReviewDate",        "dateTime": {"format": "dateOnly"}},
        {"name": "Duration",          "text": {}},
        {"name": "Tags",              "text": {"allowMultipleLines": True}},
        {"name": "FileUrl",           "hyperlink": {}},
        {"name": "FileName",          "text": {}},
        {"name": "LicenceType",       "choice": {"choices": ["Free / Open Access","Creative Commons","Company Licence","Paid — Individual","Proprietary"]}},
        {"name": "Publisher",         "text": {}},
        {"name": "Visibility",        "choice": {"choices": ["All Staff","Department Only","Management Only"]}},
        {"name": "SensitivityLevel",  "choice": {"choices": ["Standard — Internal","Confidential","Restricted"]}},
        {"name": "ReviewerNotes",     "text": {"allowMultipleLines": True}},
        {"name": "Status",            "choice": {"choices": ["Pending Review","Active","Archived","Rejected"], "defaultValue": "Pending Review"}},
        {"name": "ReferenceNumber",   "text": {}},
    ]

    for col in columns:
        try:
            graph_post(token, f"/sites/{CONFIG['SITE_ID']}/lists/{list_id}/columns", col)
            print(f"  ✓ Column added: {col['name']}")
        except Exception as e:
            print(f"  ✗ Column failed: {col['name']} — {e}")

    print(f"\n  List ID to add to CONFIG: {list_id}\n")
    return list_id


# ── RUN ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("LearnHub — SharePoint Write Access Test")
    print("========================================\n")

    token = get_token()

    # Uncomment this if you still need to find your SITE_ID:
    # discover_site_id(token,
    #     "siemensenergyag.sharepoint.com",
    #     "/teams/M_Digitial_Grid_Documents")

    CONFIG["LIST_ID"] = create_training_list(token)
    # No LIST_ID yet? Create a new test list and use its ID for this run.
    if CONFIG["LIST_ID"] == "YOUR_LIST_ID_HERE":
        CONFIG["LIST_ID"] = test_create_list(token)

    results = {}
    try:
        test_read_access(token);   results["Read access"]    = "✓ PASS"
    except Exception as e:         results["Read access"]    = f"✗ FAIL — {e}"

    try:
        test_list_write(token);    results["List write"]     = "✓ PASS"
    except Exception as e:         results["List write"]     = f"✗ FAIL — {e}"

    try:
        test_file_upload(token);   results["File upload"]    = "✓ PASS"
    except Exception as e:         results["File upload"]    = f"✗ FAIL — {e}"

    print("── Summary ──────────────────────────────────────────")
    for k, v in results.items():
        print(f"  {k:<20} {v}")
    print()