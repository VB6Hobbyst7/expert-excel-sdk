# Excel Add-in for Nano-Saas

This Excel Add-in allows users to quickly process their data using the Boon Logic Nano-Saas.

>**NOTE:** In order to use this package, it is necessary to acquire a BoonNano license from Boon Logic, Inc.  A startup email will be sent providing the details for using this package.

- __Website__: [https://github.com/boonlogic/expert-matlab-sdk](https://github.com/boonlogic/expert-matlab-sdk)
- __Documentation__: [https://github.com/boonlogic/expert-excel-sdk/Documentation](./Documentation)

--------------------------
### Loading the Excel Add-in
1. Download the `Boonnano.xlam` file from the [Boonlogic Github page](https://github.com/boonlogic/expert-excel-sdk/)
1. If it is not already added, activate the Developer tab in Excel. [For instructions on how to do this, go here](https://support.office.com/en-us/article/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)

2. From the Developer tab, select the option with a gear icon, labelled Excel Add-ins.

3. Select `Browse...` and select the `Boonnano.xlam` file downloaded in step one.

4. Make sure `Boonnano` is checked in the list and click ok.

5. To start using the Add-in, add the macro to the quick access toolbar for ease of future use. [See this site for instructions.](https://www.howtoexcel.org/tips-and-tricks/how-to-add-a-macro-to-the-quick-access-toolbar/)

6. Now, the icon in the Quick Access Toolbar is all you need to get started.

------------
### License Configuration

Note: A license must be obtained from Boon Logic to use the BoonNano Excel Add-in

The license should be saved in ~/.BoonLogic.license on unix machines or C:/Users/\<user\>/.BoonLogic.license on windows machines. This file will contain the following format:

```json
{
  "default": {
    "api-key": "API-KEY",
    "server": "WEB ADDRESS",
    "api-tenant": "API-TENANT"
  }
}
```

The *API-KEY*, *WEB ADDRESS*, and *API-TENANT* will be unique to your obtained license.

The .BoonLogic.license file will be consulted by the BoonNano Excel Add-in to successfully find and authenticate with your designated server.
>**NOTE:** It is important that the file is named and placed correctly.  
>Check for:
>  - the file starting with a period
>  - both the B and the L in BoonLogic is capitalized
>  - the extension is a .license file
