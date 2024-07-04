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