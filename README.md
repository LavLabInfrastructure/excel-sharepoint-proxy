# SharePoint Excel Proxy

Super simple Python script that acts as a proxy server to retrieve values from a specified cell in an Excel file stored in SharePoint via Microsoft Graph API.
Useful for getting excel values without a POST request, however should be in a private network as it bypasses Microsoft authentication.
Used in DicomEdit Patient ID mapping

## Configuration

Before running the script, you must set the following environment variables:

- `TENANT_ID`: The tenant ID of your Azure Active Directory.
- `SCOPE`: The scope for the Microsoft Graph API. Typically set to `https://graph.microsoft.com/.default`.
- `USERNAME`: The username of the account with access to the Excel file.
- `PASSWORD`: The password of the account with access to the Excel file.
- `CLIENT_ID`: The client ID of your registered application in Azure Active Directory.
- `CLIENT_SECRET`: The client secret of your registered application in Azure Active Directory.
- `GRAPH_DRIVE_ID`: The ID of the drive in SharePoint where the Excel file is located.
- `GRAPH_FILE_ID`: The ID of the Excel file in SharePoint.
- `GRAPH_WORKSHEET_ID`: The ID of the worksheet in the Excel file.

## Usage

### Building with Docker

1.. Build the Docker image:

   \```
   docker build -t sharepoint_excel_proxy .
   \```

3. Run the Docker container, passing the environment variables:

   \```
   docker run -d -p 8080:8080 \
   -e TENANT_ID=your_tenant_id \
   -e SCOPE=https://graph.microsoft.com/.default \
   -e USERNAME=your_username \
   -e PASSWORD=your_password \
   -e CLIENT_ID=your_client_id \
   -e CLIENT_SECRET=your_client_secret \
   -e GRAPH_DRIVE_ID=your_graph_drive_id \
   -e GRAPH_FILE_ID=your_graph_file_id \
   -e GRAPH_WORKSHEET_ID=your_graph_worksheet_id \
   --name sharepoint_excel_proxy sharepoint_excel_proxy
   \```

4. Access the proxy server as described in the Running Locally section.

### Running Locally

1. Install the required library:

   \```
   pip3 install requests
   \```

2. Set the necessary environment variables as mentioned in the Configuration section.

3. Run the script:

   \```
   python3 sharepoint_excel_proxy.py
   \```

4. The proxy server will start listening on port 8080. You can access a specific cell value by sending a GET request with the `key` parameter, where `key` is the cell address (e.g., `A1`):

   \```
   http://localhost:8080/?key=A1

## Notes

- The script uses the Resource Owner Password Credentials (ROPC) grant type to authenticate with the Microsoft Graph API. This requires the username and password of the account with access to the Excel file.
- The script caches the access token and the values retrieved from the Excel file to optimize performance.
- The script is designed to handle multiple concurrent requests using `ForkingTCPServer`.

## License

This script is provided "as is" without any warranty. Feel free to use and modify it as needed.
