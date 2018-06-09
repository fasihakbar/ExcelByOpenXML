// JavaScript source code

function callEntityAction(entityId, entityName, entityPluralName, actionUniqueName, inputParams) {
    var result = {};
    entityId = entityId.replace(/\{|\}/gi, '');
    var serverURL = Xrm.Page.context.getClientUrl();
    // pass the id as inpurt parameter
    var req = new XMLHttpRequest();
    // specify name of the entity, record id and name of the action in the Wen API Url
    req.open("POST", serverURL + "/api/data/v8.2/" + entityPluralName + "(" + entityId + ")/Microsoft.Dynamics.CRM." + actionUniqueName, false);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.onreadystatechange = function () {
        if (this.readyState == 4 /* complete */) {
            req.onreadystatechange = null;
            debugger;
            if (this.status == 204) {
                var data = null;
                if (this.response != null && this.response != undefined && this.response != '') {
                    data = JSON.parse(this.response);
                }
                else {
                    data = null;
                }
                result.IsSuccess = true;
                result.IsFailure = false;
                result.ResponseData = data;
                return result;
            }
            else if (this.status == 200) {
                debugger;
                if (this.response != null && this.response != undefined && this.response != '') {
                    var data = JSON.parse(this.response);
                    result.IsSuccess = true;
                    result.IsFailure = false;
                    result.ResponseData = data;
                    return result;
                }
            }
            else {
                debugger;
                var error = JSON.parse(this.response).error;
                result.IsSuccess = false;
                result.IsFailure = true;
                result.ResponseData = error;
                return result;
            }
        }
    };
    if (inputParams == null || inputParams == '') {
        req.send();
    }
    else {
        req.send(window.JSON.stringify(inputParams));
    }
    return result;
    // send the request with the data for the input parameter
}

function callSendExcelFileInEmailAction() {
    debugger;
    //Alert.showLoading();
    var entityId = Xrm.Page.data.entity.getId().replace(/\{|\}/gi, '');
    var entityName = 'account';
    var entityPluralName = 'accounts';
    var actionUniqueName = 'new_SendExcelFileInEmail';
    
    var inputParams = null;
    var result = callEntityAction(entityId, entityName, entityPluralName, actionUniqueName, inputParams);
    debugger;
    if (result.ResponseData != null && result.ResponseData != undefined) {
        var pluginTrace = result.ResponseData['PluginTrace'];
        var message = result.ResponseData['Message'];
        if (message == 'Success')
        {
            alert('Excel file created successfully and send to contacts in email.');
        }
        else {
            alert('There is something wrong with file creation or email sending.');
        }
    }
}