"""
This script demonstrates how to list all OneNote notebooks in a user's account.
It can use either the OAuth code flow or the client credentials flow.
Neither works for some reason related to licensing if you only have a personal account
despite this working in the Graph Explorer and despite all the Microsoft hype in docs
and by 'Co-Pilot'.

It's interesting the resource needs to be access via a different url depending on the flow used,
both resulting in different error messages.

The code flow results in a 404 error with the message

    'OneDrive for Business for this user account cannot be retrieved'

The client credentials flow results in a 404 error with them message

    'The tenant does not have a valid SharePoint license.'

One can do the same with less code using Microsoft's unsupported identity library
and the msgraph library, but that's a lot of dependencies for a couple of REST calls.

    from azure.identity import ClientSecretCredential
    from msgraph import GraphServiceClient

    token = ClientSecretCredential(
        client_id="",
        client_secret="",
        tenant_id=""
    ).get_token("https://graph.microsoft.com/.default")

    graph_client = GraphServiceClient(token.token)
    oath_data = graph_client.me.onenote.notebooks.get()

    oath_data = graph_client.me.onenote.notebooks.get()
    print(oath_data)

This uses OAuth via Azure AD to authenticate the user and obtain an access token.
Configurtion values are needed an .env file.  See the code for the required values.
"""

from collections import namedtuple
import os
import webbrowser

import msal
import requests
from flask import Flask, request
from werkzeug.serving import make_server
from requests_toolbelt.utils.dump import dump_all
from dotenv import load_dotenv


AzureAdOAuthConfig = namedtuple("AzureAdOAuthConfig",
                                   "app_client_id app_registration_secret oauth_login_url user_email_address")
login_url_base = "https://login.microsoftonline.com"
load_dotenv()
azure_ad_config = AzureAdOAuthConfig(os.getenv("AZURE_AD_APP_REGISTERATION_CLIENT_ID"),
                                     os.getenv("AZURE_AD_APP_REGISTRATION_SECRET"),
                                     f"{login_url_base}/{os.getenv('AZURE_AD_APP_REGISTRATION_TENANT_ID')}",
                                     os.getenv("USER_EMAIL_ADDRESS")
                                     )

msal_client_app = msal.ConfidentialClientApplication(
    client_id=azure_ad_config.app_client_id,
    client_credential=azure_ad_config.app_registration_secret,
    authority=azure_ad_config.oauth_login_url,
)
oath_data = None
flask_ap = Flask(__name__)

use_oath_code_flow = True

oath_scopes = ["Notes.Read.All"] if use_oath_code_flow else ["https://graph.microsoft.com/.default"]

# http://localhost:5000/oathcallback is the url registered in the azure app registration
@flask_ap.route("/oauthcallback")  #
def handle_request():
    global oath_data
    oath_data = msal_client_app.acquire_token_by_authorization_code(
        request.args["code"], scopes=oath_scopes)
    # close the browser window after consent is granted
    return "<script type=\"application/javascript\">window.close();</script>".encode("UTF-8")

if use_oath_code_flow:
    # trigger the OAuth code flow to obtain "Bearer" token
    webbrowser.get().open(msal_client_app.get_authorization_request_url(scopes=oath_scopes))
    # azure ad only allows 'localhost' and not '127.0.0.1' and per OAuth the spec must match so we config that here
    flask_app = make_server("localhost", 5000, flask_ap)
    # flask's server derives from python's own BaseServer with its handle_request method
    # a perfect means to handle a single request and exit return the thread of control to the caller
    flask_app.handle_request()
    bearer_token = oath_data["access_token"]
else: # use the client credentials flow
    oath_data = msal_client_app.acquire_token_silent(scopes=oath_scopes, account=None)
    if not oath_data:
        # logging.info("No suitable token exists in cache. Let's get a new one from Azure AD.")
        oath_data = msal_client_app.acquire_token_for_client(scopes=oath_scopes)
    # end new

# Access the token from the result
bearer_token = oath_data.get("access_token")

resource_url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
if not use_oath_code_flow:
    resource_url =  f"https://graph.microsoft.com/v1.0/users/{azure_ad_config.user_email_address}/onenote/notebooks"

r = requests.get(resource_url,
    headers={
        "Authorization": f"Bearer {bearer_token}",
        "Content-Type": "application/json",
    },
)
print(dump_all(r).decode("utf8"))