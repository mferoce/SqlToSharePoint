# SqlToSharePoint
Reads data using sql, authenticate into MS Graph, writes data in Excel in Sharepoint

Add a .properties file in your resource directory.

It should look like this:

```
app.clientId=574140a4-7a4d-4da5-9576-9ccf4bc19397
app.clientSecret=yourclientsecret
app.clientSecretId=f907d8ac-c291-4232-8aa1-9ada8cbe5875
app.tenantId=35d34c8b-1dd2-4536-b4cb-8f192379747f
app.userId=35d34c8b-1dd2-4536-b4cb-8f192379747f
app.authTenant=common
app.graphUserScopes=offline_access,Files.ReadWrite.All,User.Read
app.fileId = DD92CD78236BD745
app.sheetName = test
app.redirectUri = htpp://localhost:8080

```
