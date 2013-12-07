# Setup Salesforce Consumer Key and Secret

Salesforce Documentation:
[http://www.salesforce.com/us/developer/docs/api_rest/Content/intro_defining_remote_access_applications.htm](http://www.salesforce.com/us/developer/docs/api_rest/Content/intro_defining_remote_access_applications.htm)

1. Create developer account at [developer.salesforce.com](http://developer.salesforce.com)
2. From the Setup page for your develop account, select App Setup > Create > Apps
3. Click New in the Connected Apps section
4. If necessary, create a prefix for your developer account
5. Complete Form using the following information:

    **App Name**: Full name used to identify app (e.g. Excel Client Example)
    
    **API Name**: Unique Id for identifying app (e.g. ExcelClientExample)
    
    **Contact Email**: Email used for any app related issues

6. Check Enable OAuth Settings
7. For Callback URL, use https://login.salesforce.com/services/oauth2/callback

    (The Callback URL is not currently used in the Excel Client so the example value is used)

7. Select the minimal level of OAuth Scope that works for your application and save
8. Copy the Consumer Key and Consumer Secret for the Connected App page
