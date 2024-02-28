import pandas as pd
import requests
import json
import base64
import variable

data = pd.read_excel("TC.xlsx")
data = data[data['RESULT'] == 'Fail']
temp = ""
temoId = ""
# Azure DevOps organization URL
org_url = 'https://dev.azure.com/kloudping'

# Project name
project = 'Testing-CSV'

# Request headers
headers = {
    'Content-Type': 'application/json-patch+json',
    'Authorization': 'Basic ' + base64.b64encode(bytes(':' + variable.pat, 'ascii')).decode('ascii')
}

# Create a PBI
pbi_details = {
    'op': 'add',
    'path': '/fields/System.Title',
    'value': 'Sample PBI'
}
pbi_api_url = f'{org_url}/{project}/_apis/wit/workitems/$Product Backlog Item?api-version=6.0'
pbi_response = requests.post(pbi_api_url, headers=headers, data=json.dumps([pbi_details]))

if pbi_response.status_code == 200:
    pbi_id = pbi_response.json()["id"]
    
    print("Product Backlog Item created with ID:", pbi_id)

    for index, row in data.iterrows():
        if temp == "" or temp != row["TESTCASE SCENARIO"]:
            temp = row["TESTCASE SCENARIO"]
            bug_details = [
                {
                    "op": "add",
                    "path": "/fields/System.Title",
                    "value": temp
                },
                {
                    "op": "add",
                    "path": "/relations/-",
                    "value": {
                        "rel": "System.LinkTypes.Hierarchy-Reverse",
                        "url": f"{org_url}/{project}/_apis/wit/workitems/{pbi_id}",
                        "attributes": {
                            "name": "Parent"
                        }
                    }
                }
            ]

            bug_api_url = f'{org_url}/{project}/_apis/wit/workitems/$Bug?api-version=6.0'
            bug_response = requests.post(bug_api_url, headers=headers, data=json.dumps(bug_details))

            if bug_response.status_code != 200:
                print("Failed to create bug:", bug_response.text)
                exit()

            bug_id = bug_response.json()["id"]
            print("Bug created with ID:", bug_id)
            temoId = bug_id

        # create task using temoId
        task_details = [
            {
                "op": "add",
                "path": "/fields/System.Title",
                "value": row["EXPECTED RESULT"]
            },
            {
                "op": "add",
                "path": "/relations/-",
                "value": {
                    "rel": "System.LinkTypes.Hierarchy-Reverse",
                    "url": f"{org_url}/{project}/_apis/wit/workitems/{temoId}",
                    "attributes": {
                        "name": "Parent"
                    }
                }
            }
        ]
        task_api_url = f'{org_url}/{project}/_apis/wit/workitems/$Task?api-version=6.0'

        task_response = requests.post(task_api_url, headers=headers, data=json.dumps(task_details))

        if task_response.status_code != 200:
           print(f"Failed to create task '{row['EXPECTED RESULT']}': {task_response.text}")

        else:
            print(f"Task '{row['EXPECTED RESULT']}' created successfully under Bug:", temoId)