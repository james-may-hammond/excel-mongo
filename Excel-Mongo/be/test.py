"""
Module: test.py
Description: Simple test script for backend functionality.
Dependencies: None
"""
import requests

url = "https://127.0.0.1:8000/create_collection"
headers = {
    "Content-Type": "application/json",
    "x-mongo-uri": "mongodb://localhost",
    "x-mongo-db": "testdb"
}
data = {
    "collection": "test_collection"
}

response = requests.post(url, json=data, headers=headers, verify=False)
print("STATUS:", response.status_code)
print("BODY:", response.text)
