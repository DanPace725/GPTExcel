import azure.functions as func
import logging
import os
import msal
import requests
import json

def get_token():
    """Get a token for Microsoft Graph API."""
    client_id = os.environ["GraphClient"]
    authority = os.environ["GraphAuthority"]
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

def update_excel_sheet(token, file_endpoint, range_address, values):
    """Update an Excel sheet with given values."""
    endpoint = f"{file_endpoint}/workbook/worksheets/Sheet1/range(address=\'{range_address}\')"
    payload = {"values": values}
    response = make_graph_api_request(token, endpoint, method='PATCH', data=payload)
    return response

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)





@app.route(route="gptExcel_http_trigger")
def gptExcel_http_trigger(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
# Extract payload from the incoming request
        req_body = req.get_json()

        # Assuming the payload contains the range and values for the Excel update
        range_address = req_body.get("range")  # e.g., "A2:A3"
        values = req_body.get("values")  # e.g., [["Hello World"], ["How are you?"]]

        if not range_address or not values:
            return func.HttpResponse(
                "Invalid request: range and values are required.",
                status_code=400
            )

        token = get_token()
        # Define the values and range here, or extract them from the request
        base_path = os.environ.get("GraphDriveBasePath")

        # You might want to dynamically determine the file_endpoint based on the request or configuration
        file_endpoint = f"{base_path}/items/017IM2XLK7SDZKFIY33BDL2GZNQLRHV2OL"
        response = update_excel_sheet(token, file_endpoint, range_address, values)

        # Return the response from updating the Excel sheet
        return func.HttpResponse(
            body=json.dumps(response),
            status_code=200,
            headers={"Content-Type": "application/json"}
        )
    except ValueError:
        return func.HttpResponse(
            "Invalid JSON in request body.",
            status_code=400)

    except Exception as e:
        logging.error(f"Error: {e}")
        return func.HttpResponse(
        str(e),
        status_code=500
    )
@app.route(route="getExcelData", methods=["GET"])
def get_excel_data(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a GET request.')

    try:
        token = get_token()
        # Use an environment variable for the file endpoint
        base_path = os.environ.get("GraphDriveBasePath")
        file_endpoint = f"{base_path}/items/017IM2XLK7SDZKFIY33BDL2GZNQLRHV2OL"
        req_body = req.get_json(silent=True)

        if not req_body or "range" not in req_body:
            return func.HttpResponse(
                "Invalid request: 'range' is required in the request body.",
                status_code=400
            )

        range_address = req_body["range"]

        endpoint = f"{file_endpoint}/workbook/worksheets/Sheet1/range(address=\'{range_address}\')"
        response = make_graph_api_request(token, endpoint, method='GET')

        if response.status_code != 200:
            return func.HttpResponse(
                "Failed to retrieve data from Excel sheet.",
                status_code=response.status_code
            )

        return func.HttpResponse(
            body=json.dumps(response.json()),
            status_code=200,
            headers={"Content-Type": "application/json"}
        )

    except Exception as e:
        logging.error(f"Error: {e}")
        return func.HttpResponse(
            "An error occurred processing your request.",
            status_code=500
        )
    

@app.route(route="getDriveItems", methods=["GET"])
def get_drive_items(req: func.HttpRequest) -> func.HttpResponse:
    """Azure Function to retrieve the list of items in the root directory of a OneDrive drive and return only their names.

    Args:
        req (func.HttpRequest): The incoming HTTP request.

    Returns:
        func.HttpResponse: JSON response containing the list of item names in the drive's root.
    """
    logging.info('Processing a request to retrieve item names from the drive\'s root directory.')

    try:
        drive_base_path = os.environ["GraphDriveBasePath"]
        token = get_token()

        # Append '/root/children' to the drive_base_path to get items in the root directory
        drive_items_url = f"{drive_base_path}/root/children"
        response = make_graph_api_request(token, drive_items_url, method='GET')

        # Extract item names from the response
        item_names = [item['name'] for item in response['value']]

        # Return only the item names in the response
        return func.HttpResponse(body=json.dumps(item_names), status_code=200, headers={"Content-Type": "application/json"})
    except Exception as e:
        logging.error(f"Error: {e}")
        return func.HttpResponse(str(e), status_code=500)
