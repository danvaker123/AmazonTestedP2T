import os
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext


def download_from_sharepoint(site_url, username, password, sharepoint_file_path, local_file_path):
    """Download a file from SharePoint."""
    auth_context = AuthenticationContext(site_url)
    if not auth_context.acquire_token_for_user(username, password):
        raise Exception("Authentication failed")

    ctx = ClientContext(site_url, auth_context)
    sharepoint_file_url = f"{site_url}/{sharepoint_file_path}"
    with open(local_file_path, "wb") as local_file:
        file = ctx.web.get_file_by_server_relative_url(sharepoint_file_url)
        file.download(local_file).execute_query()
    print(f"File downloaded to {local_file_path}")


def upload_to_sharepoint(site_url, username, password, local_file_path, sharepoint_file_path):
    """Upload a file to SharePoint."""
    auth_context = AuthenticationContext(site_url)
    if not auth_context.acquire_token_for_user(username, password):
        raise Exception("Authentication failed")

    ctx = ClientContext(site_url, auth_context)
    with open(local_file_path, "rb") as local_file:
        target_folder = ctx.web.get_folder_by_server_relative_url(os.path.dirname(sharepoint_file_path))
        target_folder.upload_file(os.path.basename(sharepoint_file_path), local_file.read()).execute_query()
    print(f"File uploaded to {sharepoint_file_path}")


# Set up SharePoint credentials and paths
username = os.getenv('SHAREPOINT_USERNAME','achyut.kumar@us.gt.com')
password = os.getenv('SHAREPOINT_PASSWORD', 'Aman@magic10')
site_url = "https://gtus365.sharepoint.com/sites/technologysolutionsbangalore"
sharepoint_input_path = "/Documents/Team Shared documents/P2T-POC/input_data.xlsx"
sharepoint_output_path = "/Documents/Team Shared documents/P2T-POC/output_data.xlsx"

# Example usage:
local_input_path = r"C:\Users\US81896\OneDrive - Grant Thornton LLP\All\Desktop\input.xlsx"   # Local path to save downloaded file
local_output_path = r"C:\Users\US81896\OneDrive - Grant Thornton LLP\All\Desktop\Output.xlsx" # Local path for file to upload

# Downloading file
download_from_sharepoint(site_url, username, password, sharepoint_input_path, local_input_path)

# Uploading file
upload_to_sharepoint(site_url, username, password, local_output_path, sharepoint_output_path)
