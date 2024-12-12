import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

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

    # Ensure the correct SharePoint folder URL format
    folder_path = f"/sites/{site_url.split('/')[-1]}/Shared Documents/Testing"
    file_name = os.path.basename(local_output_path)  # Extract file name from the local path

    # Upload the file to SharePoint
    target_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
    target_file = target_folder.upload_file(file_name, file_content)
    ctx.execute_query()

    print(f"Uploaded {local_output_path} to {folder_path}")


# Main function to execute the download and upload process
def main():
    # Your credentials for Azure AD and SharePoint
    tenant_id = '16af8cf7-eeb9-4533-a377-abac8f72cc4e'  # Your Directory ID
    client_id = 'acfdceea-7647-440b-b263-756462670cee'  # Your Application ID
    client_secret = 'EC78Q~_i8liXcdR8MXLbMwmpAlEB3MA.51hzSb7W'  # Your Client Secret Value

    # SharePoint details
    site_url = "https://gtus365dev.sharepoint.com/sites/TechnologySolutionsBangalore"  # Replace with your SharePoint site URL
    sharepoint_input_path = "/Shared Documents/Testing/input_data.xlsx"
    sharepoint_output_path = "/Shared Documents/Testing"  # Folder path where file will be uploaded

    # Local paths for downloading and uploading files
    local_input_path = r"C:\Users\US81896\OneDrive - Grant Thornton LLP\All\Desktop\help.xlsx"  # Local path to save downloaded file
    local_output_path = r"C:\Users\US81896\OneDrive - Grant Thornton LLP\All\Desktop\help2.xlsx"  # Local path for file to upload

    # Download the Excel file from SharePoint
    download_from_sharepoint(site_url, client_id, client_secret, sharepoint_input_path, local_input_path)

    # Upload the output file to SharePoint
    upload_to_sharepoint(site_url, client_id, client_secret, sharepoint_output_path, local_output_path)


if __name__ == "__main__":
    main()
