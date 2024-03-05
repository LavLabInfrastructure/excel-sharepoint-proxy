#!/usr/local/bin/python3
import socketserver
from http.server import SimpleHTTPRequestHandler
from typing import Any
import requests
from urllib.parse import urljoin, urlparse, parse_qs
from io import BytesIO
from os import environ
from time import time
import logging
logging.basicConfig(level=logging.WARNING)
PORT = 8080

# https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token
IDP_URL = "https://login.microsoftonline.com"
TENANT_ID = environ["TENANT_ID"]
TOKEN_ENDPOINT = "/oauth2/v2.0/token"

GRANT_TYPE = "password"
SCOPE = environ["SCOPE"]
USERNAME = environ["USERNAME"]
PASSWORD = environ["PASSWORD"]
CLIENT_ID = environ["CLIENT_ID"]
CLIENT_SECRET = environ["CLIENT_SECRET"]

# https://graph.microsoft.com/v1.0/
GRAPH_ROOT_URL = "https://graph.microsoft.com"
GRAPH_ROOT_ENDPOINT = "/v1.0"

# https://graph.microsoft.com/v1.0/drives/{GRAPH_DRIVE_ID}
GRAPH_DRIVE_ENDPOINT = "/drives"
GRAPH_DRIVE_ID = "/" + environ["GRAPH_DRIVE_ID"]


# https://graph.microsoft.com/v1.0/drives/{GRAPH_DRIVE_ID}/items/GRAPH_ITEM_ID
GRAPH_FILE_ENDPOINT="/items"
GRAPH_FILE_ID = "/" + environ["GRAPH_FILE_ID"]

# https://graph.microsoft.com/v1.0/drives/{GRAPH_DRIVE_ID}/items/GRAPH_ITEM_ID ...
# /workbook/worksheets/
GRAPH_WORKSHEETS_ENDPOINT = "/workbook/worksheets"
GRAPH_WORKSHEET_ID = "/" + environ["GRAPH_WORKSHEET_ID"]

class SharePointExcelProxy(SimpleHTTPRequestHandler):
    # TODO should really put this in a constructor
    token_url = urljoin(IDP_URL, TENANT_ID+TOKEN_ENDPOINT)
    token_body = {
        'grant_type': GRANT_TYPE,
        'scope': SCOPE,
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'username': USERNAME,
        'password': PASSWORD
    }
    
    mapper_url = urljoin(GRAPH_ROOT_URL, GRAPH_ROOT_ENDPOINT + 
                        GRAPH_DRIVE_ENDPOINT + GRAPH_DRIVE_ID +
                        GRAPH_FILE_ENDPOINT + GRAPH_FILE_ID +
                        GRAPH_WORKSHEETS_ENDPOINT + GRAPH_WORKSHEET_ID)
    logging.info(mapper_url)
    savedMap = {}
    access_token = None
    token_expires_at = None

    def getAccessToken(self):
        """post resource owner creds at graph.microsoft.com to get access token"""
        if self.access_token is None or self.token_expires_at is None or self.token_expires_at <= time():
            logging.info("getting new access token!")
            response = requests.post(self.token_url, data=self.token_body).json()
            self.access_token = response["access_token"]
            self.token_expires_at = time() + response["expires_in"] - 60  # expire 1 minute early
        return self.access_token
    
    
    def getMSGraphText(self, key):
        logging.info(self.mapper_url + f"/range(address='{key}'")
        rv = requests.get(
                self.mapper_url + f"/range(address='{key}')",
                headers = {"Authorization": f"Bearer {self.getAccessToken()}"}
                ).json() 
        try: 
            rv = rv["text"]
        except KeyError: 
            logging.error(rv)
            return None
        
        while type(rv) is list:
            rv = rv[0]

        return rv
    

    def getMapping(self,key):
        try: 
            return self.savedMap[key]
        
        except KeyError:
            logging.warning(f"No mapping saved for key: {key}")
            val = self.getMSGraphText(key)

            if val is not None:
                self.savedMap.update({key:val})

            return val


    def do_GET(self): 
        logging.info(f"Recieved GET request from: {self.client_address[0]}")

        key = parse_qs(urlparse(self.path).query)["key"][0]
        rv = self.getMapping(key)

        if rv is None:
            # logging.error(f"Sending 404 to client: {self.client_address[0]} for key: {key}")
            # self.send_response(404)
            logging.error(f'Key: {key} was not valid, returning the key as the value.')
            rv = key
        
        self.send_response(200)
        self.end_headers()
        self.copyfile( BytesIO(rv.encode()), self.wfile)
    def log_message(self, format: str, *args: Any) -> None:
        return None

if __name__ == "__main__":
    httpd = socketserver.ThreadingTCPServer(('', PORT), SharePointExcelProxy)
    logging.info("Now serving on port: " + str(PORT))
    httpd.serve_forever()