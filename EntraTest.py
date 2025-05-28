import msal
import requests
import json
import os
import openpyxl
import sys

CLIENT_ID = os.environ.get('ENTRACLIENTID')
TENANT_ID = os.environ.get('ENTRATENANTID')
CLIENT_SECRET = os.environ.get('ENTRACLIENTSECRET')
SCOPES = ["https://graph.microsoft.com/.default"]
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
GRAPH_API_USERS_URL = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled"
GRAPH_API_GROUPS_URL = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName"

xlssPath = sys.argv[1]
targetGroups = sys.argv[2:]

def parse_xlsx_users(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active  # Get the active sheet
        xlsx_users = []
        for row in sheet.iter_rows(): # Iterate through each row
            xlsx_users.extend([[cell.value for cell in row]])  # Get values from each cell
        return xlsx_users
    except Exception as e:
        print("Unexpected error:", e)

def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if 'access_token' in result: return result['access_token']
    else:
        print(f"Error getting token: {result.get('error_description')}")
        return None

def fetch_users(access_token):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(GRAPH_API_USERS_URL, headers=headers)
    if response.status_code == 200:
        users = response.json()
        return users.get('value', [])
    else:
        print(f"Error fetching users: {response.status_code} - {response.text}")
        return []
    
def fetch_groups(access_token):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(GRAPH_API_GROUPS_URL, headers=headers)
    if response.status_code == 200:
        users = response.json()
        return users.get('value', [])
    else:
        print(f"Error fetching groups: {response.status_code} - {response.text}")
        return []
    
def fetch_groupMembers(access_token, groupID):
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f"https://graph.microsoft.com/v1.0/groups/{groupID}/members?$select=id,displayName,userPrincipalName"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        users = response.json()
        return users.get('value', [])
    else:
        print(f"Error fetching groupMembers for group ({groupID}): {response.status_code} - {response.text}")
        return []
    
def add_user_to_group(access_token, userID, groupID):
    headers = {'Authorization': f'Bearer {access_token}','Content-Type': 'application/json'}
    url = f"https://graph.microsoft.com/v1.0/groups/{groupID}/members/$ref"
    body = {"@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{userID}"}
    response = requests.post(url, headers=headers, json=body)
    if response.status_code == 204:
        print(f"Successfully added user ({userID}) to group ({groupID}).")
        return True
    else:
        print(f"Error adding user to group ({groupID}): {response.status_code} - {response.text}")
        return False

def main():
    access_token = get_access_token()
    xlsx_users = parse_xlsx_users(xlssPath)[1:]
    print(f"parsed_xlss_users: {xlsx_users}")

    if access_token:
        users = fetch_users(access_token)
        # disabled_users = [user for user in users if not user.get('accountEnabled', True)]
        groups = fetch_groups(access_token)

        if users and groups:
            print(f"groups: {groups}")
            for targetGroupName in targetGroups:
                groupID = checkGroupExists(groups, targetGroupName)
                if(groupID != -1): 
                    print(f"\nAdding to {targetGroupName}:")
                    for xlsx_user in xlsx_users: addUserToGroup(access_token, users, xlsx_user, targetGroupName, groupID)
                else: print(f"\n{targetGroupName}: exists in Entra: False")
        else: print("Either users or groups not found.")
    else:
        print("Failed to obtain access token.")

def checkGroupExists(groups, groupName):
    for group in groups:
        if(group.get('displayName')==groupName): return group.get('id')
    return -1

def addUserToGroup(access_token, users, xlsx_user, targetGroupName, groupID):
    groupMembers = fetch_groupMembers(access_token, groupID)
    if(userExists(users, xlsx_user[1])!=-1):
        if groupMemberExists(groupMembers, xlsx_user[1]):
            print(f"\"{xlsx_user[1]}\" already exists in {targetGroupName}")
        else:
            print(f"\"{xlsx_user[1]}\" does not exist in {targetGroupName}")
            add_user_to_group(access_token, getUserID(users, xlsx_user[1]), groupID)
    else: print(f"\"{xlsx_user[1]}\" does not exist in users")

def userExists(users, userName):
    for user in users:
        if(user.get('displayName')==userName): return user.get('id')
    else: return -1

def groupMemberExists(groupMembers, memberName):
    for member in groupMembers:
        if(member.get('displayName')==memberName): return True
    return False

def getUserID(users,userName):
    for user in users:
        if(user.get('displayName')==userName): return user.get('id')
    return -1

if __name__ == '__main__':
    main()
