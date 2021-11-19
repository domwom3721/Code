from carto.auth import APIKeyAuthClient

USERNAME="boweryres"
USR_BASE_URL = "https://{user}.carto.com/".format(user=USERNAME)
auth_client = APIKeyAuthClient(api_key="665cd4b5ebe654dad7101dcf0048718f45282835", base_url=USR_BASE_URL)

from carto.datasets import DatasetManager
#local file or url
local_file_or_URL = ""

dataset_manager = DatasetManager(auth_client)
dataset = dataset_manager.create(LOCAL_FILE_OR_URL)
