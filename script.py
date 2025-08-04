import requests as rq
import json as js
import os as system
from dotenv import load_dotenv;
import pandas as pd

# Initializing variables
load_dotenv();
token = system.environ['TOKEN'] # Your zephyr token
excel_document = system.environ['DOCUMENT_TEST_CASES'] # The route of you excel document, needs to have two columns (ID Test Case, Folder)
if not token or not excel_document:
    raise ValueError("Not all the variables are set")

base_url: str = 'https://api.zephyrscale.smartbear.com/v2'

headers: dict = {
    "Authorization": f'Bearer {token}',
    "Content-Type": "application/json" 
}

# Getting a test case
def get_test_case(test_case_id):
    request = f'{base_url}/testcases/{test_case_id}'
    response = rq.get(request, headers=headers)
    response = response.json()
    print("Got the test case")
    return response
    # print(response.status_code)
    # print(js.dumps(response.json(), indent=4))

# Move the test case to another folder
def move_test_cases(folder_id: int, test_case_id: str):
    test_case = get_test_case(test_case_id=test_case_id)
    test_case['folder']['id'] = folder_id
    request = f'{base_url}/testcases/{test_case_id}'
    converting_to_json = js.dumps(test_case)
    # print(converting_to_json)
    response = rq.put(request, headers=headers, data=converting_to_json)
    response.raise_for_status()
    print("Test case moved")
    # print(js.dumps(response.json(), indent=4))

# Get the folder information
def get_folder(folder_id: int):
    request = f'{base_url}/folders/{folder_id}'
    response = rq.get(request, headers=headers)
    response = response.json()
    print(response)
    return response

# Get the folder id by test case
def get_folder_by_test_case(test_case_id: str):
    response = get_test_case(test_case_id)
    folder = response["folder"]["id"]
    print(f'Folder: {folder}')
    return folder

# Getting data
def read_excel(route):
    excel_document = pd.read_excel(route, header=0)
    # print(excel_document)
    # dp = pd.DataFrame(data=excel_document, columns=["ID_Test_Case", "ID_Folder"])
    print(excel_document.count())
    confirm = str(input("This is what you're going to move, are you sure you're going to do it? [Y/N] \n"))
    confirm.lower()
    if confirm == "y":
        test_cases: dict = excel_document.to_dict()
        print(test_cases)
        for index in range(len(test_cases)):
            _test_case = test_cases["ID Test Case"][index]
            _folder = test_cases["Folder"][index]
            print(f'-- Test case {_test_case} to folder {_folder} --')

            if _test_case != None and _folder != None:
                move_test_cases(folder_id=_folder, test_case_id=_test_case)
            else:
                print(f"The test case with ID {_test_case} or folder {_folder} doesn't have all values, review the document!!")
                continue
    else:
        print("You exited")

# Start
read_excel(excel_document)