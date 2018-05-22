
Register the application to use Graph:
1. Sign into the Application Registration Portal  (https://apps.dev.microsoft.com) using "work or school account".

![Sign in to your tenant](images/sign in to your tenant.png)

2. Sign into your tenant (i.e. <user>@<tenant>.onmicrosoft.com, <password>)

![My applications](images/my applications.png)

3. Click on the "Add and app" button.

4. Enter the name for the app as "Proseware Tasks", and choose Create application.

![Register your app](images/register your app.png)

5. The registration page displays, listing the properties of your app.

6. IMPORTANT: Copy the Application Id and save it somewhere. This is the unique identifier for your app. You'll use this value to configure your app.
	
![Copy the AppId](images/copy the appid.png)

7. Under Platforms, choose Add Platform.
	
![Add a Platform](images/add a platform.png)
8. Choose Web.
	
![Choose Web Platform](images/choose web platform.png)

9. Make sure the Allow Implicit Flow check box is selected, and enter https://localhost:44382/Home.html as the Redirect URI.
10. In Microsoft Graph Permissions, next to Delegated Permissions, click "Add"
	
11. Add the permissions so they match the following: 
	
12. Choose Save.


