import urllib.parse
import urllib.request
import json

def get_onedrive_token(code: str):

    apl_client_id= "a330b0e9-f2f1-472b-bb2b-fcfb9f3767f8"
    client_secret = "zZy8Q~tSIfmdJT6C30vBEBJGlpIcra39OWwWlcor"
    redirect_url = "http://localhost:4200/"

    url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    method = "POST"
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    params = {
        "client_id": apl_client_id
        ,"redirect_uri": redirect_url
        ,"client_secret": client_secret
        ,"code": code
        ,"grant_type": "authorization_code"
    }

    encoded_param = urllib.parse.urlencode(params).encode()

    request = urllib.request.Request(url, data=encoded_param, method=method, headers=headers)
    with urllib.request.urlopen(request) as res:
        body = res.read()
        dat = json.loads(body)
        print(dat)

if __name__ == '__main__':

    code = "aaaaaaaaaaaaaaaaaaaaaaaa"
    get_onedrive_token(code)

