# Proof of Concept

This proof of concept demonstrates how to store an Excel file to Onedrive/Sharepoint and open the file with Excel Online using Dotnet Framework 4.61.

## Microsoft Graph Api

This application use the Microsoft Graph API to store files in a Sharepoint Site. [Documentation for the Microsoft Graph SDK](https://learn.microsoft.com/en-us/graph/sdks/sdks-overview).

## Authentication

The web application authenticates to Microsoft using a Service Principal.

### Create a Service Principal

To create a service princial go to the [Azure Portal](https://portal.azure.com) and sign in as a Global Administrator

1. Go to **Microsoft Entra Id**
2. Click **Enterprise Applications**
3. Click **New application**
4. Click **Create your own application**
5. Give the aplication a name and select *Register an application to integrate with Microsoft Entra ID (App you're developing)*, click **Next**
6. Select *Accounts in this organizational directory only (`Your tenanat name`  only - Single tenant)* and click **Register**

#### Grant Graph API permissions to the Service Principal

1. Search for your application in **Entra Id | App Registrations**
2. Select **Api Permissions** and click **Add a permission**
3. Select **Microsoft Graph**
4. Select **Application Permission**
5. Type *Files* in the serachbar and select `Files.ReadWrite.All` and click **Add Permission**
6. Click **Grant admin consent for `Your tenanat name`**

#### Create a Client Secret for your service principal

1. Search for your application in **Entra Id | App Registrations**
2. Select **Certificates & secrets** and click **New Client Secret**
3. Give your secret a name and expire date and click **Add**

>[!IMPORTANT]
>Copy and store the client secret value in a secure location. This is the only time you are allowed to do this

#### Configure the application to use the service principal

The application use the [Azure.Identity](https://www.nuget.org/packages/Azure.Identity) nuget package, and the code expects the following environment variables

| Variable name         | Value                                          |
|:--------------------  |:---------------------------------------------- |
| `AZURE_CLIENT_ID`     | ID of a Microsoft Entra application            |
| `AZURE_CLIENT_SECRET` | The application's client secrets               |
| `AZURE_TENANT_ID`     | ID of the application's Microsoft Entra tenant |
