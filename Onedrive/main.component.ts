oneDriveConnect(){
    const target_url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" 
    + "client_id="a330b0e9-f2f1-472b-bb2b-fcfb9f3767f8"
    + "&scope=offline_access%20files.readwrite.all"
    + "&response_type=code"
    + "&redirect_uri=http://localhost:4200/";

    location.href=target_url;
  }
