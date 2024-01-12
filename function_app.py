import azure.functions as func
import logging
import os
import msal

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

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="gptExcel_http_trigger")
def gptExcel_http_trigger(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
        # Attempt to get the access token
        token = get_token()
        logging.info("Successfully obtained access token.")
    except Exception as e:
        logging.error(f"Error obtaining access token: {e}")
        return func.HttpResponse(
            "Error obtaining access token.",
            status_code=500
        )

    # Existing logic to handle the request
    name = req.params.get('name')
    if not name:
        try:
            req_body = req.get_json()
        except ValueError:
            pass
        else:
            name = req_body.get('name')

    if name:
        return func.HttpResponse(f"Hello, {name}. This HTTP triggered function executed successfully.")
    else:
        return func.HttpResponse(
            "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.",
            status_code=200
        )

# Add any additional functionality as needed
