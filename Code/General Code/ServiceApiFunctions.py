import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retrydef 
import json

def UpdateServiceDb(type, csv_name, csv_path):
    if type == None:
        return
    print(f'Updating service database: {type}')

    url = f'http://localhost:5300/api/v1/update/{type}'
    dropbox_path = f'https://www.dropbox.com/home/Research/Market Analysis/Market/{csv_name}'
    payload = { 'location': dropbox_path }

    retry_strategy = Retry(
        total=3,
        status_forcelist=[400, 404, 409, 500, 503, 504],
        allowed_methods=["POST"],
        backoff_factor=5
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    http = requests.Session()
    http.mount("https://", adapter)
    http.mount("http://", adapter)

    response = http.post(url, json=payload)
    if response.status_code == 200:
        # Delete the temporary CSV
        os.remove(csv_path)
        print('Service successfully updated')
