# Prints displayName and webUrl for every list on the configured SharePoint site.
from learnhub_sharepoint_test import get_token, graph_get, CONFIG

token = get_token()
lists = graph_get(token, f"/sites/{CONFIG['SITE_ID']}/lists")
for l in lists["value"]:
    print(l["displayName"], "->", l["webUrl"])
