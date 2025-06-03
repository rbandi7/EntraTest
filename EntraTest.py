import msal
import requests
import json
import os
import openpyxl
import sys


    
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
