CLIENT_ID = os.environ.get('ENTRACLIENTID')Add commentMore actions
TENANT_ID = os.environ.get('ENTRATENANTID')
CLIENT_SECRET = os.environ.get('ENTRACLIENTSECRET')
SCOPES = ["https://graph.microsoft.com/.default"]
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'

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
    url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        users = response.json()
        return users.get('value', [])
    else:
        print(f"Error fetching users: {response.status_code} - {response.text}")
        return []
    
def fetch_groups(access_token):
    headers = {'Authorization': f'Bearer {access_token}'}
    url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName"
    response = requests.get(url, headers=headers)
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

def remove_user_from_group(access_token, userID, groupID):
    headers = {'Authorization': f'Bearer {access_token}','Content-Type': 'application/json'}
    url = f"https://graph.microsoft.com/v1.0/groups/{groupID}/members/{userID}/$ref"
    response = requests.delete(url, headers=headers)
    if response.status_code == 204:
        print(f"Successfully added user ({userID}) to group ({groupID}).")
        return True
    else:
        print(f"Error adding user to group ({groupID}): {response.status_code} - {response.text}")
        return False
        
def enable_user_account(access_token, userID):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
    data = {"accountEnabled": True}
    response = requests.patch(url, headers=headers, json=data)
    if response.status_code == 204:
        print(f"Successfully re-enabled user ({userID})")
        return True
    else:
        print(f"Error re-enabling user ({userID})")
        return False

def disable_user_account(access_token,userID):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
    data = {"accountEnabled": False}
    response = requests.patch(url, headers=headers, json=data)
    if response.status_code == 204:
        print(f"Successfully disabled user ({userID})")
        return True
    else:
        print(f"Error disabling user ({userID})")
        return False

def get_auth_methods(access_token, user_id):
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/authentication/methods"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        methods = response.json().get('value', [])
        usable = []
        non_usable = []
        for method in methods:
            if method['@odata.type'] == "#microsoft.graph.emailAuthenticationMethod": usable.append(method)
            else: non_usable.append(method)
        return usable, non_usable
    else:
        print("Failed to get auth methods:", response.status_code, response.text)
        return [], []

def remove_email_from_nonusable(access_token, non_usable_methods):
    headers = {'Authorization': f'Bearer {access_token}'}
    removed = False
    for method in non_usable_methods:
        if method['@odata.type'] == "#microsoft.graph.emailAuthenticationMethod":
            method_id = method['id']
            user_id = method['id'].split('/')[0]  # Might need adjustment based on actual ID format
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}/authentication/emailMethods/{method_id}"
            response = requests.delete(url, headers=headers)
            if response.status_code == 204:
                print("Removed email method:", method_id)
                removed = True
            else:
                print("Failed to remove:", response.status_code, response.text)
    return removed

def add_email_to_usable(access_token, user_id, email_address):
    headers = {'Authorization': f'Bearer {access_token}','Content-Type': 'application/json'}
    data = {"emailAddress": email_address}
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/authentication/emailMethods"
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 201:
        print("Successfully added usable email method.")
        return True
    else:
        print("Failed to add email method:", response.status_code, response.text)
        return False
