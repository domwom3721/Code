from carto.auth import APIKeyAuthClient

USERNAME="boweryres"
USR_BASE_URL = "https://{user}.carto.com/".format(user=USERNAME)
auth_client = APIKeyAuthClient(api_key="665cd4b5ebe654dad7101dcf0048718f45282835", base_url=USR_BASE_URL)
