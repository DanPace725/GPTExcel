GPT Excel Backend

Overview
This project is a Python-based Azure Function App designed to facilitate real-time interaction between an OpenAI GPT model and Excel files stored in OneDrive. It acts as a backend service that allows the GPT model to retrieve, update, and manipulate data within Excel sheets using the Microsoft Graph API. By exposing a set of RESTful API endpoints, this backend enables seamless communication between the GPT model and Excel files, allowing for dynamic updates, data retrieval, and real-time analysis of spreadsheet data based on GPT-generated responses. The project is deployed to Azure Functions, ensuring scalable and secure processing for real-time Excel interactions.

Features
Authentication: Uses MSAL (Microsoft Authentication Library) for secure token acquisition to interact with Microsoft Graph API.
Excel File Operations: Provides endpoints to list Excel files, fetch data from Excel sheets, and update cell values in Excel files using the Graph API.
HTTP Triggers: RESTful APIs for accessing and manipulating Excel files via HTTP requests.
Automated Deployment: Utilizes GitHub Actions to build and deploy to Azure Functions.

Deployment
The project uses GitHub Actions for Continuous Integration and Continuous Deployment (CI/CD). The workflow (.github/workflows/main_gptexcelbackend.yml) defines the steps to build and deploy the function app to Azure Functions.

Steps:
Checkout Code: GitHub Actions checks out the repository.
Set Up Python Environment: Installs Python 3.11 and sets up a virtual environment.
Install Dependencies: Installs the required packages listed in requirements.txt.
Package and Deploy: Zips the project and deploys it to Azure Functions.
Setup and Configuration
Prerequisites
Azure Function App: Create an Azure Function App in your Azure subscription.
Azure Active Directory: Register an application in Azure AD to enable MSAL-based authentication.
Environment Variables
GraphClient: The client ID of your Azure AD application.
GraphAuthority: The Azure AD tenant authority URL.
GraphSecret: The client secret for your application.
GraphDriveBasePath: The base path of your OneDrive drive to access files.
Local Development
Install Dependencies:

bash
Copy code
pip install -r requirements.txt
Run Locally: Use VS Code's integrated terminal and the provided tasks for local development:

bash
Copy code
func start
Testing: Test endpoints locally using tools like Postman or curl.

Notes
Azure Function Version: This project uses Azure Functions runtime version 4.0.
Python Version: Python 3.11 is specified in the GitHub workflow.
MSAL Authentication: The app uses Microsoft Graph API for interacting with Excel files, and authentication is managed via the MSAL library.
License
This project is licensed under the MIT License. See the LICENSE file for more details.
