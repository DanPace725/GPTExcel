import requests
import json
import logging
import msal

def load_config(file_path):
    """Load configuration from a file."""
    with open(file_path, 'r') as file:
        return json.load(file)

def get_token(config):
    """Get a token for Microsoft Graph API."""
    app = msal.ConfidentialClientApplication(
            config["client_id"], authority=config["authority"],
            client_credential=config["secret"])
    
    result = app.acquire_token_silent(config["scope"], account=None)
    print(app)
    if not result:
            logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
            result = app.acquire_token_for_client(scopes=config["scope"])
            
            
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Error acquiring token: {result}")
        

def make_graph_api_request(token, endpoint, method='GET', data=None):
    """Make a request to the Microsoft Graph API."""
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    if method == "GET":
        response = requests.get(endpoint, headers=headers)
    elif method == "POST":
        response = requests.post(endpoint, headers=headers, json=data)
    elif method == "PUT":     
        response = requests.put(endpoint, headers=headers, json=data)
    elif method == "PATCH":     
        response = requests.patch(endpoint, headers=headers, json=data)
    # Add more methods as needed
    else:
        raise ValueError("HTTP method not supported")

    return response.json()

def update_excel_sheet(token, endpoint, values):
    """Update an Excel sheet with given values."""
    payload = {"values": values}
    response = make_graph_api_request(token, endpoint, method='PATCH', data=payload)
    return response

def main():
    config = load_config("gpt_excel/parameters.json")
    token = get_token(config)
    

    """ values = [
        ["Hello World"],["How are you?"]
        ]
    
    range = "/range(address=\'A2:A3\')"
    response = update_excel_sheet(token, config["example_file_endpoint"] + range, values)
    print(response)

    if response.get('error', None):
        print("Failed to update excel sheet.")
    else:
        print("Success")"""


if __name__ == "__main__":
    main()
