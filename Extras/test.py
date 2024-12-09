import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential


# Function to get access token from Azure AD
def get_access_token(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    body = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'  # This scope is for accessing Microsoft Graph
    }

    response = requests.post(url, data=body)

    if response.status_code == 200:
        access_token = response.json().get('access_token')
        print("Access token obtained successfully!")
        # Return the access token directly as a string
        return access_token
    else:
        print(f"Failed to obtain access token: {response.status_code}")
        print(response.text)
        return None


# Function to download file from SharePoint
def download_from_sharepoint(site_url, client_id, client_secret, sharepoint_input_path, local_input_path):
    # Create a context with the client credentials
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

    # Get the file and download it
    file = ctx.web.get_file_by_server_relative_url(sharepoint_input_path)
    with open(local_input_path, 'wb') as local_file:
        file.download(local_file).execute_query()

    print(f"Downloaded {sharepoint_input_path} to {local_input_path}")


# Function to upload file to SharePoint
def upload_to_sharepoint(site_url, client_id, client_secret, sharepoint_output_path, local_output_path):
    # Create a context with the client credentials
    ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

    # Read the local file to upload
    with open(local_output_path, 'rb') as local_file:
        file_content = local_file.read()

    # Upload the file to SharePoint
    target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_output_path)
    target_file = target_folder.upload_file(local_output_path.split("\\")[-1], file_content)
    ctx.execute_query()

    print(f"Uploaded {local_output_path} to {sharepoint_output_path}")


# Main function to execute the download and upload process
def main():
    # Your credentials for Azure AD and SharePoint
    tenant_id = '7d76d45a-a201-4a68-bf3a-597f0a5fa533'  # Your Directory ID
    client_id = 'b25e7795-c312-4adf-b5ce-c391ef549d76'  # Your Application ID
    client_secret = 'oM68Q~YvnnTh6nxwESnQ0kG_Qlr6UR2c6GjMXb5t'  # Your Client Secret Value

    # SharePoint details
    site_url = "https://gtus365.sharepoint.com/sites/technologysolutionsbangalore"  # Replace with your SharePoint site URL
    sharepoint_input_path = "/Shared Documents/Team Shared documents/P2T- POC/Input and Output File  Folder/input_data.xlsx"
    sharepoint_output_path = "/Shared Documents/Team Shared documents/P2T- POC/Input and Output File  Folder/output_data.xlsx"

    # Local paths for downloading and uploading files
    local_input_path = r"C:\Users\US81896\OneDrive - Grant Thornton LLP\All\Desktop\Input file.xlsx"  # Local path to save downloaded file
    local_output_path = r"C:\Users\US81896\OneDrive - Grant Thornton LLP\All\Desktop\Output file.xlsx"  # Local path for file to upload

    # Get access token from Azure AD (Note: This step is not needed if using ClientCredential directly)
    access_token = get_access_token(tenant_id, client_id, client_secret)

    if access_token:
        # Download the Excel file from SharePoint
        download_from_sharepoint(site_url, client_id, client_secret, sharepoint_input_path, local_input_path)

        # Upload the output file to SharePoint
        upload_to_sharepoint(site_url, client_id, client_secret, sharepoint_output_path, local_output_path)


if __name__ == "__main__":
    main()