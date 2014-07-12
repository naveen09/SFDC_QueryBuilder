/**
 * @author      Naveen_Kumar
 * @description SOQL Query Builder for Google Spreadsheets
 */
/**
 * Key of UserProperties for Salesforce Username.
 * @type {String}
 * @const
 */
var USERNAME_PROPERTY_NAME = "username";

/**
 * Key of UserProperties for Salesforce Password.
 * @type {String}
 * @const
 */
var PASSWORD_PROPERTY_NAME = "password";

/**
 * Key of UserProperties for Salesforce Security Token.
 * @type {String}
 * @const
 */
var SECURITY_TOKEN_PROPERTY_NAME = "securityToken";

/**
 * Key of UserProperties for Salesforce Session Id.
 * @type {String}
 * @const
 */
var SESSION_ID_PROPERTY_NAME = "sessionId";

/**
 * Key of UserProperties for serviceUrl.
 * @type {String}
 * @const
 */
var SERVICE_URL_PROPERTY_NAME = "serviceUrl";


/**
 * Key of UserProperties for instance url.
 * @type {String}
 * @const
 */
var INSTANCE_URL_PROPERTY_NAME = "instanceUrl";

/**
 * Key of UserProperties for sandbox url.
 * @type {String}
 * @const
 */
var IS_SANDBOX_PROPERTY_NAME = "isSandbox";

/**
 * Key of UserProperties for next records url.
 * @type {String}
 * @const
 */
var SOBJECT_ATTRIBUTES_PROPERTY_NAME = "sObjectAttributes";

var SANDBOX_SOAP_URL = "https://test.salesforce.com/services/Soap/u/29.0";

var PRODUCTION_SOAP_URL = "https://www.salesforce.com/services/Soap/u/29.0";

var currentData = new Array();
/**
 * @return String Username.
 */
function getUsername() {
    var key = UserProperties.getProperty(USERNAME_PROPERTY_NAME);
  
    if (key == null) {
        key = "";
    }
    return key;
}

/**
 * @param String Username.
 */
function setUsername(key) {
    UserProperties.setProperty(USERNAME_PROPERTY_NAME, key);
}

/**
 * @return String Password.
 */
function getPassword() {
    var key = UserProperties.getProperty(PASSWORD_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
}

/**
 * @param String Password.
 */
function setPassword(key) {
 UserProperties.setProperty(PASSWORD_PROPERTY_NAME, key);
}

/**
 * @return String Security Token.
 */
function getSecurityToken() {
    var key = UserProperties.getProperty(SECURITY_TOKEN_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
}

/**
 * @param String Security Token.
 */
function setSecurityToken(key) {
    UserProperties.setProperty(SECURITY_TOKEN_PROPERTY_NAME, key);
}

/**
 * @return String Session Id.
 */
function getSessionId() {
    var key = UserProperties.getProperty(SESSION_ID_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
}

/**
 * @param String Session Id.
 */
function setSessionId(key) {
    UserProperties.setProperty(SESSION_ID_PROPERTY_NAME, key);
}

/**
 * @return String Instance URL.
 */
function getInstanceUrl() {
    var key = UserProperties.getProperty(INSTANCE_URL_PROPERTY_NAME);
    if (key == null) {
        key = "";
    }
    return key;
}

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
    UserProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
}


/**
 * @param String use sandbox url.
 */
function setUseSandbox(key) {
    UserProperties.setProperty(IS_SANDBOX_PROPERTY_NAME, key);
}

/**
 * @return bool if using sandbox.
 */
function getUseSandbox() {
    var key = UserProperties.getProperty(IS_SANDBOX_PROPERTY_NAME);
    if (key == null) {
        key = false;
    }
    return key;
}

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
    UserProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
}

/**
 * @return bool if using sandbox.
 */
function getSfdcSoapEndpoint() {
    var isSandbox = getUseSandbox() == "true" ? true : false;
    if (isSandbox)
        return SANDBOX_SOAP_URL;
    else
        return PRODUCTION_SOAP_URL;
}

function getRestEndpoint() {
    //Move this logic to the property
    var queryEndpoint = ".salesforce.com";

    var endpoint = getInstanceUrl().replace("api-", "").match("https://[a-z0-9]*");

    return endpoint + queryEndpoint;
}



function onInstall() {
    onOpen();
}
/* create menu and menu items here */
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{
        name: "Create/Run Query",
        functionName: "renderQueryDialog"
    }];

    ss.addMenu("SFDC Connector", menuEntries);
   if(credentialsAvailable()){
     renderSettingsDialog();
   }
}

/** Retrieve config params from the UI and store them. */
function saveConfiguration(e) {
    setUsername(e.parameter.username);
    setPassword(e.parameter.password);
    setSecurityToken(e.parameter.securityToken);
    setUseSandbox(e.parameter.sandbox);
    login();
    var app = UiApp.getActiveApplication();
    app.close();
    return app;
}
/* Load settings dialog to enter username, password and security token issued by Force.com */
function renderSettingsDialog() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Salesforce Configuration");
    app.setStyleAttribute("padding", "10px");

    var helpLabel = app.createLabel("Enter your Username, Password, and Security Token");
    helpLabel.setStyleAttribute("text-align", "justify");

    var usernameLabel = app.createLabel("Username:");
    var username = app.createTextBox();
    username.setName("username");
    username.setWidth("75%");
    username.setText(getUsername());

    var passwordLabel = app.createLabel("Password:");
    var password = app.createPasswordTextBox();
    password.setName("password");
    password.setWidth("75%");
    password.setText(getPassword());

    var securityTokenLabel = app.createLabel("Security Token:");
    var securityToken = app.createTextBox();
    securityToken.setName("securityToken");
    securityToken.setWidth("75%");
    securityToken.setText(getSecurityToken());

    var sandboxLabel = app.createLabel("Sandbox:");
    var sandbox = app.createCheckBox();
    sandbox.setName("sandbox");
    sandbox.setValue(getUseSandbox() == "true" ? true : false);

    var saveHandler = app.createServerClickHandler("saveConfiguration");
    var saveButton = app.createButton("Login", saveHandler);

    var listPanel = app.createGrid(4, 2);
    listPanel.setStyleAttribute("margin-top", "10px")
    listPanel.setWidth("100%");
    listPanel.setWidget(0, 0, usernameLabel);
    listPanel.setWidget(0, 1, username);
    listPanel.setWidget(1, 0, passwordLabel);
    listPanel.setWidget(1, 1, password);
    listPanel.setWidget(2, 0, securityTokenLabel);
    listPanel.setWidget(2, 1, securityToken);
    listPanel.setWidget(3, 0, sandboxLabel);
    listPanel.setWidget(3, 1, sandbox);

    // Ensure that all form fields get sent along to the handler
    saveHandler.addCallbackElement(listPanel);

    var dialogPanel = app.createFlowPanel();
    dialogPanel.add(helpLabel);
    dialogPanel.add(listPanel);
    dialogPanel.add(saveButton);
    app.add(dialogPanel);
    doc.show(app);
}
/* perform login with SOAP envelop*/
function login() {

    var message = "<?xml version='1.0' encoding='utf-8'?>" + "<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' " + "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://" + "www.w3.org/2001/XMLSchema'>" + "<soap:Body>" + "<login xmlns='urn:partner.soap.sforce.com'>" + "<username>" + getUsername() + "</username>" + "<password>" + getPassword() + getSecurityToken() + "</password>" + "</login>" + "</soap:Body>" + "</soap:Envelope>";

  Logger.log(getSfdcSoapEndpoint());
    var httpheaders = {
        SOAPAction: "login"
    };
    var parameters = {
        method: "POST",
        contentType: "text/xml",
        headers: httpheaders,
        payload: message
    };

    try {
        var result = UrlFetchApp.fetch(getSfdcSoapEndpoint(), parameters).getContentText();
        var soapResult = Xml.parse(result, false);

        setSessionId(soapResult.Envelope.Body.loginResponse.result.sessionId.getText());
        setInstanceUrl(soapResult.Envelope.Body.loginResponse.result.serverUrl.getText());
        Browser.msgBox("Log-in Success...");
    } catch (e) {
        Logger.log("EXCEPTION!!!");
        Logger.log(e);
      if(!credentialsAvailable()){
        Browser.msgBox(e);
      }else{
        renderSettingsDialog();
      }
    }

}
/* Load spreadsheet with data*/
function renderGridData(object, renderHeaders) {
    var sheet = SpreadsheetApp.getActiveSheet();

    var data = [];
    var sObjectAttributes = {};

    //Need to always build headers for row length/rendering
    var headers = buildHeaders(object.records);

    if (renderHeaders) {
        data.push(headers);
    }

    for (var i in object.records) {
        var values = [];
        for (var j in object.records[i]) {
            if (j != "attributes") {
                values.push(object.records[i][j]);
            } else {
                var id = object.records[i][j].url.substr(object.records[i][j].url.length - 18, 18);
                //Logger.log(id);
                sObjectAttributes[id] = object.records[i][j].type;
            }
        }
        data.push(values);
    }

    //Logger.log(sheet.getLastRow());
    var destinationRange = sheet.getRange(sheet.getLastRow() + 1, 1, data.length, headers.length);
    destinationRange.setValues(data);
}

/* build spreadsheet headers */
function buildHeaders(records) {
    var headers = [];
    for (var i in records[0]) {
        if (i != "attributes")
            headers.push(i);
    }
    //Logger.log(headers);
    return headers;
}
/* Execute query and populate sheet */
function sendSoqlQuery(e) {
    var app = UiApp.getActiveApplication();
    var query = e.parameter.soql;
    populateSheetUsingQuery(query);
    return app;
}

function processResults(results) {
    var object = Utilities.jsonParse(results);
    return object;
}


/* render Query builder dialog */
function renderQueryDialog() {
    var w = 800;
    var h = 450;
    var panelHeight = 300;
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Query Builder");

    app.setHeight(h);
    app.setWidth(w);

    UserProperties.setProperty("selectclause", "");
    UserProperties.setProperty("whereclause", "");

    currentData.length = 0;
    UserProperties.setProperty("SOCurrentData", Utilities.jsonStringify(currentData));

    var dialogPanel = app.createTabPanel();

    var queriesTab = app.createVerticalPanel().setWidth(w - 100).setHeight(h - 100).setId("querytab");
    var builderTab = app.createVerticalPanel().setWidth(w - 100).setHeight(h - 100);
    var advancedTab = app.createVerticalPanel().setWidth(w - 100).setHeight(h - 100);
    var historyTab = app.createVerticalPanel().setWidth(w - 100).setHeight(h - 100).setId("historytab");

    var mainContainer = app.createGrid(1, 3).setWidth("100%");
    var leftContainer = app.createGrid(3, 1).setWidth("100%").setId("leftContainer");
    var rightContainer = app.createGrid(5, 1).setWidth("100%");

    queriesTab.setStyleAttribute("background", "#F0F0F0");
    builderTab.setStyleAttribute("background", "#F0F0F0");
    advancedTab.setStyleAttribute("background", "#F0F0F0");
    historyTab.setStyleAttribute("background", "#F0F0F0");

    mainContainer.setStyleAttribute("padding", "10px");
    mainContainer.setStyleAttribute("background", "#F0F0F0");


    mainContainer.setWidget(0, 0, leftContainer);
    mainContainer.setWidget(0, 1, rightContainer);

    var objectLabel = app.createLabel("Object:");
    var objectCombo = app.createListBox().setName("sfobject").setTitle("Select Salesforce Object").addItem("");
    
    var sdescribe = Utilities.jsonParse(fetch(getRestEndpoint() + "/services/data/v29.0/sobjects/"));
    
    var OBJS = sdescribe.sobjects;
  
    for (j in OBJS) {
          var fld = OBJS[j].name;
          objectCombo.addItem(fld);
     }
  
    var sfobjecthandler = app.createServerHandler('showSelectedSFObject').addCallbackElement(leftContainer);
    objectCombo.addChangeHandler(sfobjecthandler)

    var objectHP = app.createGrid(1, 3).setPixelSize(100, 30);


    objectHP.setWidget(0, 0, objectLabel);
    objectHP.setWidget(0, 1, objectCombo);

    leftContainer.setWidget(0, 0, objectHP);

    var selectFieldLabel = app.createLabel("Select Field(s)");
    selectFieldLabel.setStyleAttribute("width", "85px")

    var sortBtn = app.createButton("Sort").setTitle("Click to sort fields");
    sortBtn.setStyleAttribute("background", "#0F9D58");
    sortBtn.setStyleAttribute("color", "#FFFFFF");

    var selectFieldChecbox = app.createCheckBox("Select All").setId("selectallbtn").setName("selectallbtn");
    selectFieldChecbox.setStyleAttribute("width", "85px")

    var fieldHP = app.createGrid(1, 3);
    fieldHP.setWidth("100%");

    fieldHP.setWidget(0, 0, selectFieldLabel);
    fieldHP.setWidget(0, 1, sortBtn);
    fieldHP.setWidget(0, 2, selectFieldChecbox);
    leftContainer.setWidget(1, 0, fieldHP);


    var fieldContainer = app.createListBox(true).setId("field-Container").setName("fieldContainer");
    fieldContainer.setWidth(300);
    fieldContainer.setHeight(200);
    fieldContainer.setStyleAttribute("background", "#FFFFFF");
    fieldContainer.setStyleAttribute("border", "1px solid #444444");
    leftContainer.setWidget(2, 0, fieldContainer);
    UserProperties.setProperty("fieldContainercounter", 0);

    var selectAllHandler = app.createServerHandler("selectallfields").addCallbackElement(fieldContainer);
    selectFieldChecbox.addClickHandler(selectAllHandler);


    var sortHandler = app.createServerHandler('_sortItems').addCallbackElement(fieldContainer);
    sortBtn.addClickHandler(sortHandler);

    var selectionHandler = app.createServerHandler('updateSelectClause').addCallbackElement(fieldContainer);
    fieldContainer.addChangeHandler(selectionHandler);

    /* end of left container */

    /*start of right container*/
    var filterLabel = app.createLabel("Filters:");

    var filterScrollPanel = app.createScrollPanel().setPixelSize(w / 2, 200);
    var filterContainer = app.createGrid(2, 1);

    var table = app.createFlexTable().setId("table").setTag('0');
    UserProperties.setProperty("totalCombos", 0);
    filterScrollPanel.setStyleAttribute("border", "1px solid #444444");
    filterScrollPanel.setStyleAttribute("background", "#FFFFFF");

    var filterHeaderPanel = app.createHorizontalPanel().setWidth(300);

    var fieldLabel = app.createLabel("Field");
    var conditionLabel = app.createLabel("Condition");
    var valueLabel = app.createLabel("Value");

    filterHeaderPanel.add(fieldLabel);
    filterHeaderPanel.add(conditionLabel);
    filterHeaderPanel.add(valueLabel);

    filterContainer.setWidget(0, 0, filterHeaderPanel);
    table.setWidth(300);
    filterContainer.setWidget(1, 0, table);

    //add combos here.
    addMemberCombos(app);

    filterScrollPanel.add(filterContainer);

    rightContainer.setWidget(0, 0, filterLabel);
    rightContainer.setWidget(1, 0, filterScrollPanel);


    var helpLabel = app.createLabel("SOQL Query");
    helpLabel.setStyleAttribute("text-align", "justify");
    var soql = app.createTextArea().setName("soql").setWidth("400px").setId("soql");


    rightContainer.setWidget(2, 0, helpLabel);
    rightContainer.setWidget(3, 0, soql);


    var btnbar = app.createGrid(1, 5);


    var sendHandler = app.createServerClickHandler("sendSoqlQuery");

    var saveQueryHandler = app.createServerClickHandler("savequery");
    saveQueryHandler.addCallbackElement(rightContainer);
    var saveQueryBtn = app.createButton("Save", saveQueryHandler).setStyleAttribute("background", "#51A351").setStyleAttribute("color", "white").setTitle("Save this query");

    var clearQueryHandler = app.createServerHandler("clearQuery").addCallbackElement(rightContainer);
    var clearBtn = app.createButton("Clear", clearQueryHandler).setStyleAttribute("background", "#C53C36").setStyleAttribute("color", "white").setTitle("Clear");;


    var sendButton = app.createButton("Run", sendHandler).setStyleAttribute("background", "#0044CC").setStyleAttribute("color", "white").setTitle("Execute query");

    var savestatus = app.createLabel("Status: ").setStyleAttribute("font-weight", "bold");
    var status = app.createLabel("-").setId("status");

    btnbar.setWidget(0, 0, saveQueryBtn);
    btnbar.setWidget(0, 1, clearBtn);
    btnbar.setWidget(0, 2, sendButton);
    btnbar.setWidget(0, 3, savestatus);
    btnbar.setWidget(0, 4, status);

    rightContainer.setWidget(4, 0, btnbar);
    sendHandler.addCallbackElement(rightContainer);
    /* end of right container*/

    builderTab.add(mainContainer);

    /* build advanced tab here */
    var advanceGrid = app.createGrid(2, 1).setWidth("100%");
    var btnbar = app.createGrid(1, 3);

    var runQueryadvHandler = app.createServerHandler("runQueryAdvMode").addCallbackElement(advanceGrid);
    var saveQueryadvHandler = app.createServerHandler("saveQueryAdvMode").addCallbackElement(advanceGrid);
    var clearQueryadvHandler = app.createServerHandler("clearQueryAdvMode").addCallbackElement(advanceGrid);

    var queryarea = app.createTextArea().setId("advancequery").setName("advancequery").setWidth(w - 150).setHeight(h - 200);
    var advanceRunBtn = app.createButton("Run query", runQueryadvHandler).setId("advancequerybtn").setStyleAttribute("background", "#0044CC").setStyleAttribute("color", "white").setTitle("Execute query");
    var clearQueryAdvanceBtn = app.createButton("Clear", clearQueryadvHandler).setId("clearQueryAdvanceBtn").setStyleAttribute("background", "#C53C36").setStyleAttribute("color", "white").setTitle("Clear");
    var saveQueryAdvanceBtn = app.createButton("Save", saveQueryadvHandler).setId("saveQueryAdvanceBtn").setStyleAttribute("background", "#51A351").setStyleAttribute("color", "white").setTitle("Save this query");


    btnbar.setWidget(0, 0, advanceRunBtn);
    btnbar.setWidget(0, 1, saveQueryAdvanceBtn);
    btnbar.setWidget(0, 2, clearQueryAdvanceBtn);

    advanceGrid.setWidget(0, 0, queryarea);
    advanceGrid.setWidget(1, 0, btnbar);

    advancedTab.add(advanceGrid);
    /* end of build advanced tab here */

    dialogPanel.add(builderTab, 'Create Query').add(queriesTab, 'Saved Queries').add(advancedTab, 'Advanced Mode').add(historyTab, "History");
    dialogPanel.selectTab(0);
    app.add(dialogPanel);

    buildQueryGrid(app);

    populateHistory();

    doc.show(app);

}
/* select all fields */
function selectallfields(e) {
    //Logger.log("inside selectallfields");
    var checked = e.parameter.selectallbtn;
    var app = UiApp.getActiveApplication();
    var fieldContainer = app.getElementById("field-Container");
    var currentData = Utilities.jsonParse(UserProperties.getProperty("SOCurrentData"));

    if (currentData.length == 0) {
        return;
    }
    fieldContainer.clear();
    var qf = "";
    for (j in currentData) {

        qf += currentData[j];
        if (j != currentData.length - 1) {
            qf += ", ";
        }
        fieldContainer.addItem(currentData[j]);
        if ("true" == checked) {
            fieldContainer.setItemSelected(parseInt(j), true);
        }
    }
    if ("true" == checked) {
        selectQueryHelper(qf);
    } else {
        selectQueryHelper("");
    }
    return app;
}
/* Run Query advanced mode*/
function runQueryAdvMode(e) {
    //Logger.log("inside runQueryAdvMode");
    var app = UiApp.getActiveApplication();
    var query = e.parameter.advancequery;
    if ("" == query) {
        return;
    } else {
        populateSheetUsingQuery(query);
    }
    return app;
}
/* Clear Query advanced mode*/
function clearQueryAdvMode(e) {
    //Logger.log("inside clearQueryAdvMode");
    var app = UiApp.getActiveApplication();
    app.getElementById("advancequery").setText("");
    return app;
}
/*advanced mode save query into script properties using this method*/
function saveQueryAdvMode(e) {
    //Logger.log("saveQueryAdvMode");
    var app = UiApp.getActiveApplication();
    var savedQueries = Utilities.jsonParse(UserProperties.getProperty("savedqueries"));

    if (null == savedQueries) {
        savedQueries = new Array();
    }
    var qry = e.parameter.advancequery;
    if ("" != qry) {
        savedQueries.push(e.parameter.advancequery);
        UserProperties.setProperty("savedqueries", Utilities.jsonStringify(savedQueries));
        buildQueryGrid(app);
        Browser.msgBox("Successfully saved the query...");
    } else {
        Browser.msgBox("Failed to save the query...");
    }
    return app;
}

/* Clear query text area */
function clearQuery(e) {
    //Logger.log("inside clearQuery");
    var app = UiApp.getActiveApplication();
    app.getElementById("soql").setText("");
    return app;
}
/* save query into script properties using this method*/
function savequery(e) {
    //Logger.log("savequery");
    var app = UiApp.getActiveApplication();
    var savedQueries;
    if (null == UserProperties.getProperty("savedqueries")) {
        savedQueries = new Array();
    } else {
        savedQueries = Utilities.jsonParse(UserProperties.getProperty("savedqueries"));
    }
    if (null == savedQueries) {
        savedQueries = new Array();
    }
    var qry = e.parameter.soql;
    if ("" != qry) {
        savedQueries.push(e.parameter.soql);
        UserProperties.setProperty("savedqueries", Utilities.jsonStringify(savedQueries));
        buildQueryGrid(app);
        app.getElementById("status").setText("Successfully saved the query...").setStyleAttribute("color", "#0F9D58");
    } else {
        app.getElementById("status").setText("Query field is empty...").setStyleAttribute("color", "#F29513");
    }
    return app;
}
/* Build query list grid in 2nd tab */
function buildQueryGrid(app) {
    //Logger.log("inside buildQueryGrid");
    var savedQueries;
    if (null == UserProperties.getProperty("savedqueries")) {
        savedQueries = new Array();
    } else {
        savedQueries = Utilities.jsonParse(UserProperties.getProperty("savedqueries"));
    }
    var rows = savedQueries.length;

    var querytab = app.getElementById("querytab");
    querytab.clear();

    var querytablscroller = app.createScrollPanel().setWidth(700).setHeight(350);
    querytab.add(querytablscroller);

    var queryGrid = app.createGrid(rows, 3).setId("querygrid").setWidth("100%");
    querytablscroller.add(queryGrid);

    var loadqueryHandler = app.createServerHandler("loadqueryfromgrid").addCallbackElement(queryGrid);
    var deletequeryHandler = app.createServerHandler("deleteQueryfromgrid").addCallbackElement(queryGrid);

    for (var i = 0; i < rows; i++) {
        var queryString = app.createLabel(savedQueries[i], true).setId("query" + i).setStyleAttribute("color", "blue");
        var loadQueryBtn = app.createButton("Run Query", loadqueryHandler).setId("loadquery" + i).setStyleAttribute("background", "#57AF57").setStyleAttribute("color", "white").setTitle("Run this query");
        var deleteQueryBtn = app.createButton(" X ", deletequeryHandler).setId("deletequery" + i).setStyleAttribute("background", "#C53C36").setStyleAttribute("color", "white").setTitle("Remove from list");
        queryGrid.setWidget(i, 0, queryString);
        queryGrid.setWidget(i, 1, loadQueryBtn);
        queryGrid.setWidget(i, 2, deleteQueryBtn);
    }

}
/* remove query from query grid */
function deleteQueryfromgrid(e) {
    //Logger.log("inside deleteQueryfromgrid");
    
    var app = UiApp.getActiveApplication();
    var savedQueries = Utilities.jsonParse(UserProperties.getProperty("savedqueries"));
    var source = e.parameter.source;
    var index = source.substring(source.length - 1);
    if (index == 0) {
        savedQueries.length = 0;
    } else {
        savedQueries.splice(index, index);
    }
    UserProperties.setProperty("savedqueries", Utilities.jsonStringify(savedQueries));
    buildQueryGrid(app);
    return app;
}
/* run query which is selected in the query grid */
function loadqueryfromgrid(e) {
    //Logger.log("loadqueryfromgrid");
    var app = UiApp.getActiveApplication();
    var source = e.parameter.source;
    var index = source.substring(source.length - 1);
    var savedQueries = Utilities.jsonParse(UserProperties.getProperty("savedqueries"));
    var queryField = savedQueries[index];

    populateSheetUsingQuery(queryField);
    return app;
}

/* display data on to the sheet after query ran successfully*/
function populateSheetUsingQuery(queryField) {
    var app = UiApp.getActiveApplication();
    var history;
    if (null == UserProperties.getProperty("history")) {
        history = new Array();
    } else {
        history = Utilities.jsonParse(UserProperties.getProperty("history"));
    }
    history.push(getDate() + "$$" + queryField);
    UserProperties.setProperty("history", Utilities.jsonStringify(history));
    populateHistory();
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.clear();
    var results = query(encodeURIComponent(queryField));
    renderGridData(processResults(results), true);
    app.close();
}
/* populate history grid */
function populateHistory() {
    //Logger.log("inside populateHistory");
    var app = UiApp.getActiveApplication();

    if (null != UserProperties.getProperty("history")) {
        var history = Utilities.jsonParse(UserProperties.getProperty("history"));

        var historytab = app.getElementById("historytab");
        historytab.clear();
        var historyscroller = app.createScrollPanel().setWidth(700).setHeight(300);
        historytab.add(historyscroller);

        var rows = history.length;
        var historyGrid = app.createGrid(rows, 2).setId("historygrid").setWidth("100%");

        var clearHistoryHandler = app.createServerHandler("clearHistory").addCallbackElement(historyGrid);
        var clearHistoryBtn = app.createButton("Clear History", clearHistoryHandler).setId("historyclearBtn");

        historyscroller.add(historyGrid);
        historytab.add(clearHistoryBtn);

        for (var i = 0; i < history.length; i++) {
            var hQuery = new Array();
            hQuery = history[i].split("$$");
            var dateLabel = app.createLabel(hQuery[0]).setStyleAttribute("font-weight", "bold");
            var queryLabel = app.createLabel(hQuery[1]).setStyleAttribute("color", "grey");
            historyGrid.setWidget(i, 0, dateLabel);
            historyGrid.setWidget(i, 1, queryLabel);
        }
    }
}

/* clear query history which is in history tab */
function clearHistory(e) {
    //Logger.log("inside clearHistory");
    var app = UiApp.getActiveApplication();
    app.getElementById("historygrid").clear();
    if (null != UserProperties.getProperty("history")) {
        var history = Utilities.jsonParse(UserProperties.getProperty("history"));
        history.length = 0;
        UserProperties.setProperty("history", Utilities.jsonStringify(history));
    }
    return app;
}

/* get current date (MM/DD/YY HH:mm:ss), used while building history */
function getDate() {
    var currentDate = new Date();
    var day = currentDate.getDate();
    var month = currentDate.getMonth() + 1;
    var year = currentDate.getFullYear();
    var hh = currentDate.getHours();
    var mm = currentDate.getMinutes();
    var ss = currentDate.getSeconds();
    return "" + month + "/" + day + "/" + year + "  " + hh + ":" + mm + ":" + ss + "";
}

/* update select clause in text area */
function updateSelectClause(e) {
    //Logger.log("inside updateSelectClause");
    var app = UiApp.getActiveApplication();
    selectQueryHelper(e.parameter.fieldContainer);
    return app;
}
/* helper function to populate query in text area */
function selectQueryHelper(fieldString) {
    var query;
    var app = UiApp.getActiveApplication();
    var queryField = app.getElementById("soql");
    queryField.setText("");
    if (fieldString != "") {
        query = "SELECT ";
        query += fieldString;
        var currentSFO = UserProperties.getProperty("currentSFObject");
        query += " FROM " + currentSFO;
        UserProperties.setProperty("selectclause", query);
        queryField.setText(query + UserProperties.getProperty("whereclause"))
    }
}
/* sort items in field list */
function _sortItems(e) {
    //Logger.log("Inside _sortItems");
    var app = UiApp.getActiveApplication();
    app.getElementById("selectallbtn").setValue(false);
    selectQueryHelper("");
    currentData = Utilities.jsonParse(UserProperties.getProperty("SOCurrentData"));
    currentData.sort();
    var fieldContainer = app.getElementById("field-Container");
    fieldContainer.clear();
    for (j in currentData) {
        fieldContainer.addItem(currentData[j]);
    }
    UserProperties.setProperty("SOCurrentData", Utilities.jsonStringify(currentData));
    updateCombos(app);
    return app;
}


/* handler for object drop down selection */
function showSelectedSFObject(e) {
    //Logger.log("inside showSelectedSFObject");

    var app = UiApp.getActiveApplication();
    var lstsize = UserProperties.getProperty("fieldContainercounter");
    var fieldContainer = app.getElementById("field-Container");
    var queryField = app.getElementById("soql");

    queryField.setText("");

    if (lstsize != 0) {
        fieldContainer.clear();
    }

    var sobject = e.parameter.sfobject;
    UserProperties.setProperty("currentSFObject", sobject);
    currentData.length = 0;

    if ("" == sobject) {
        UserProperties.setProperty("fieldContainercounter", 0);
        UserProperties.setProperty("selectclause", "");
        UserProperties.setProperty("whereclause", "");
        UserProperties.setProperty("SOCurrentData", "");
        updateCombos(app);
        return app;
    } else {
        var sdescribe = Utilities.jsonParse(fetch(getRestEndpoint() + "/services/data/v29.0/sobjects/" + sobject + "/describe/"));
        var sFields = sdescribe.fields;
        UserProperties.setProperty("fieldContainercounter", sFields.length);
        for (j in sFields) {
            var fld = sFields[j].name;
            fieldContainer.addItem(fld);
            currentData.push(fld);
        }
    }

    UserProperties.setProperty("SOCurrentData", Utilities.jsonStringify(currentData));
    updateCombos(app);
    return app;
}

/* update all combos here based on object selection */
function updateCombos(app) {
    //Logger.log("Inside updateCombos");

    var soarray = Utilities.jsonParse(UserProperties.getProperty("SOCurrentData"));
    var tag = Utilities.jsonParse(UserProperties.getProperty("totalCombos"));

    for (var i = 1; i <= tag; i++) {
        var combos = app.getElementById("fieldCombo" + i);
        combos.clear();
        combos.addItem("");
        for (j in soarray) {
            combos.addItem(soarray[j]);
        }
    }
}
/* add UI combo rows each time when add(+) button is selected */
function addMemberCombos(app) {
    //Logger.log("Inside addMeMberCombos");

    var table = app.getElementById('table');
    var tag = parseInt(table.getTag());
    var numRows = tag + 1;
    var fieldCombohandler = app.createServerHandler('updateWhereClause').addCallbackElement(table);

    if (numRows > 1) {
        table.removeCell(numRows - 1, 4);
        table.removeCell(numRows - 1, 3);
        var andorcombo = app.createListBox().setId('andorcombo' + tag).setName('andorcombo' + tag);
        andorcombo.addItem("");
        andorcombo.addItem("AND");
        andorcombo.addItem("OR");
        table.setWidget(numRows - 1, 3, andorcombo);
        andorcombo.addChangeHandler(fieldCombohandler);
    }

    var fieldCombo = app.createListBox().setId('fieldCombo' + numRows).setName('fieldCombo' + numRows).addItem("");
    var conditionCombo = app.createListBox().setId('conditionCombo' + numRows).setName('conditionCombo' + numRows).addItem("");
    var valueText = app.createTextBox().setId('valueText' + numRows).setName('valueText' + numRows);

    var soarray = Utilities.jsonParse(UserProperties.getProperty("SOCurrentData"));
    for (j in soarray) {
        fieldCombo.addItem(soarray[j]);
    }

    conditionCombo.addItem("=");
    conditionCombo.addItem("!=");
    conditionCombo.addItem("<");
    conditionCombo.addItem(">");
    conditionCombo.addItem("<=");
    conditionCombo.addItem(">=");
    conditionCombo.addItem("like");
    conditionCombo.addItem("starts with");
    conditionCombo.addItem("contains");
    conditionCombo.addItem("IS NULL");

    fieldCombo.setWidth("80px");
    conditionCombo.setWidth("80px");
    valueText.setWidth("80px");

    table.setWidget(numRows, 0, fieldCombo);
    table.setWidget(numRows, 1, conditionCombo);
    table.setWidget(numRows, 2, valueText);
    table.setTag(numRows.toString());
    UserProperties.setProperty("totalCombos", numRows);


    fieldCombo.addChangeHandler(fieldCombohandler);
    conditionCombo.addChangeHandler(fieldCombohandler);
    valueText.addChangeHandler(fieldCombohandler);

    addButtons(app);
}
/* update where clause for query building */
function updateWhereClause(e) {
    //Logger.log("updateWhereClause");
   var app = UiApp.getActiveApplication();
    var queryField = app.getElementById("soql");
    queryField.clear();
    var finalQuery = UserProperties.getProperty("selectclause");

    var tag = parseInt(e.parameter.table_tag);
    var whereClause = " WHERE ";
    for (var i = 1; i <= tag; i++) {

        if (e.parameter['andorcombo' + (i - 1)] != undefined) {
            whereClause += e.parameter['andorcombo' + (i - 1)] + " ";
        }
        whereClause += e.parameter['fieldCombo' + i] + " ";
        var condition = e.parameter['conditionCombo' + i];
        if ("starts with" == condition) {
            whereClause += "like ";
            if ("" != e.parameter['valueText' + i]) {
                whereClause += "'" + e.parameter['valueText' + i] + "%' ";
            }
        } else if ("contains" == condition) {
            whereClause += "like ";
            if ("" != e.parameter['valueText' + i]) {
                whereClause += "'%" + e.parameter['valueText' + i] + "%' ";
            }
        } else {
            whereClause += condition+" ";
            if ("" != e.parameter['valueText' + i]) {
                whereClause += "'" + e.parameter['valueText' + i] + "' ";
            }
        }
    }
    UserProperties.setProperty("whereclause", whereClause);

    finalQuery += whereClause;
    queryField.setText(finalQuery);

    return app;
}
/* add  + and -  button building */
function addButtons(app) {

    var table = app.getElementById('table');
    var numRows = parseInt(table.getTag());

    //Create handler to add/remove row
    var addRemoveRowHandler = app.createServerHandler('_addRemoveRow');
    addRemoveRowHandler.addCallbackElement(table);

    //Add row button and handler
    var addRowBtn = app.createButton('+').setId('addOne').setTitle('Add row');
    table.setWidget(numRows, 3, addRowBtn);
    addRowBtn.addMouseUpHandler(addRemoveRowHandler);

    //remove row button and handler
    var removeRowBtn = app.createButton('-').setId('removeOne').setTitle('Remove row');
    table.setWidget(numRows, 4, removeRowBtn);
    removeRowBtn.addMouseUpHandler(addRemoveRowHandler);
}

/* add each row and arraneg the UI */
function _addRemoveRow(e) {
   // Logger.log("Inside _addRemoveRow");
    var app = UiApp.getActiveApplication();
    var table = app.getElementById('table');
    var tag = parseInt(e.parameter.table_tag);
    var source = e.parameter.source;

    if (source == 'addOne') {
        table.setTag(tag.toString());
        UserProperties.setProperty("totalCombos", tag);
        addMemberCombos(app);
    } else if (source == 'removeOne') {
        if (tag > 1) {
            //Decrement the tag by one
            var numRows = tag - 1;
            table.removeRow(tag);
            //Set the new tag of the table
            table.setTag(numRows.toString());
            UserProperties.setProperty("totalCombos", numRows);
            //Add buttons in previous row
            addButtons(app);
        }
    }
    return app;
}
/**
 * @param String SQOL query
 */
function query(soql) {
    return fetch(getRestEndpoint() + "/services/data/v29.0/" + "query?q=" + soql);
}


/**
 * @param String url to fetch from SFDC via REST API
 */
function fetch(url) {

    var httpheaders = {
        Authorization: "OAuth " + getSessionId()
    };
    var parameters = {
        headers: httpheaders
    };
    try {
        return UrlFetchApp.fetch(url, parameters).getContentText();
    } catch (e) {
        Logger.log(e);
        if(!credentialsAvailable()){
        Browser.msgBox(e);
      }else{
        renderSettingsDialog();
      }
    }

}

function credentialsAvailable(){
 return (getUsername()=="" || getSecurityToken()==""||getPassword()==""); 
}