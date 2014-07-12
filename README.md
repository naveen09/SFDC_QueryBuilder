SFDC_QueryBuilder
=================

This is a Query Builder application, which is built using Google App Script (GAS) which connects to Salesforce. One can perform series of CRUD operations.

Set Up
---------

1. Goto http://script.google.com
2. create new project, add a new file code.gs (if it does not exits)
3. replace the content of code.gs with sfdcconnector.gs as shown below


```
function myFunction(){

}
```

```
var USERNAME_PROPERTY_NAME = "username";


function onInstall() {
    onOpen();
}
/* create menu and menu items here */
function onOpen() {
   ............
    }];
}

function saveConfiguration(e) {
 	...........
}
.... so on complete script
```

4. Save your project and give a name to it.
5. Goto Publish --> Deploy as web app.
6. Save it as a new version. You`ll get a confirmation message saying "This project is deployed as a web app"
7. Once done, goto http://docs.google.com and create a new spread sheet.
8. You can see "SFDC Connector" as a menu item and a modal dialog "Salesforce Configuration" will be displayed.
9. Fill in your credentials and security token (SFDC security token) and click "Login".
10. If you avoid step 9 and try to "Create" or "Run" any query, you will not be allowed to do that, you`ll be shown "Authorization Required" (the general app authorization dialog from google side).
11. Accept the "Permission" and continue.
12. Refresh the page and try to login with "Configuration" dialog. Once done. You`ll be able to pull data and build your custom query.



License
-------

    Copyright 2014 Naveen Kumar Aechan Vinod
    
    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at
    
    http://www.apache.org/licenses/LICENSE-2.0
    
    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.