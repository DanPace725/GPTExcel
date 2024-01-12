import azure.functions as func
import logging
import os
import msal
import requests

def get_token():
    """Get a token for Microsoft Graph API."""
    client_id = os.environ["GraphClient"]
    authority = "https://login.microsoftonline.com/e7ff886e-c3fe-451f-a8e2-b4e879043d56"
    client_secret = os.environ["GraphSecret"]

    app = msal.ConfidentialClientApplication(
        client_id, authority=authority,
        client_credential=client_secret
    )

    scope = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_silent(scope, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=scope)
    
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
    else:
        raise ValueError("HTTP method not supported")

    return response.json()

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="gptExcel_http_trigger")
def gptExcel_http_trigger(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
        # Obtain the access token
        token = get_token()

        # Example usage of make_graph_api_request (modify as needed)
        endpoint = "https://graph.microsoft.com/v1.0/drives/b!ddqahrDq6Eu1NVZhhGP4GtgprDkU-NJPuvcgW0p_hVC2MRe0e6t6Q63vrJkVhhG2"
        graph_response = make_graph_api_request(token, endpoint)

        # Return the Graph API response to the user
        return func.HttpResponse(
            body=str(graph_response),
            status_code=200,
            headers={"Content-Type": "application/json"}
        )

    except Exception as e:
        logging.error(f"Error: {e}")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

# Add any additional functionality as needed
