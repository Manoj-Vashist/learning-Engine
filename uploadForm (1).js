//declare global variables here
var formMode, itemId, callingSubmitFunctionFlag = 0, lastItemIdOfLib, legacyInfo,
    consumerFileFullPath, nativeFileFullPath, existingProductType,
    siteURL = _spPageContextInfo.webAbsoluteUrl,
    masterLibraryName = "eSource Content Library",
    offlineTypeListName = "RFGeSourceOfflineTypes",
    contentTypeListName = "RFGeSourceContentTypes",
    SPKMGroup = "Knowledge Management", SPCDGroup = "Curriculum Developer",
    siteServerURL = _spPageContextInfo.webServerRelativeUrl,
    isConsumer, isNative, nativeDocumentId, consumerDocumentId, waitDialog = null,
    releaseDateAlert_Today = "Release Date cannot be lower than today.",
    releaseDateAlert_Retention = "Retention Date must be later than Release Date.",
    retentionDateAlert_Today = "Retention Date cannot be lower than today.",
    retentionDateAlert_Release = "Retention Date must be later than Release Date.",
    fileNameMessage = "The uploaded file name is different than the existing file name.",
    fileExtentionMessage = "The uploaded file name has different file extension than the already uploaded file.";

$(document).ready(function () {
    var tempitemId = getParameterValues('ID');
    tempitemId = tempitemId === undefined ? getParameterValues('ItemID') : tempitemId;
    itemId = tempitemId == undefined ? "" : parseInt(tempitemId);
    tempitemId = itemId;
    checkPageMode();
    if (formMode === "newForm") {
        getLastItemIdFromLib();
        if (itemId === "") {
            initiatePopUpControl();
        }
    }
    // processTermSets function needs these js files ('sp.js', 'SP.ClientContext') before calling itself
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
        var scriptbase = _spPageContextInfo.webServerRelativeUrl + "/_layouts/15/";
        $.getScript(scriptbase + "SP.Taxonomy.js", processTermSets);
    });
    $("#accordion").accordion({ heightStyle: "Content" });
    $('.accordion-content').show();
    // allowing toggling of Accordion only in Edit mode
    if (formMode == "editForm") {
        $('.accordion-toggle').on('click', function (event) {
            event.preventDefault();
            // create accordion variables
            var accordion = $(this);
            var accordionContent = accordion.next('.accordion-content');
            // toggle accordion link open class
            accordion.toggleClass("open");
            // toggle accordion content
            accordionContent.slideToggle(250);
        });
        // checking if any Metadata property is changed in Edit mode excluding the file input controls
        $("#controlsDiv :input").change(function (evt) {
            var controlId = evt.target.id;
            if (controlId !== "NativeFileControlId" && controlId !== "ConsumerFileControlId") {
                $("#controlsDiv").data("changed", true);
            }
        });
        checkFileUploadValidation();
    }
    checkDateValidation();
    getOfflineTypes();

});

function checkDateValidation() {
    //Release Date Validation where it should always be greater than or equal to today and cannot be greater than Retention Date
    $('#ReleaseDate').datepicker({
        onSelect: function () {
            var releaseDate = $(this).datepicker('getDate');
            var retentionDate = $('#RetentionDate').datepicker('getDate');
            var todayDate = new Date();
            todayDate.setHours(0, 0, 0, 0);
            if (releaseDate < todayDate) {
                alert(releaseDateAlert_Today);
                $('#ReleaseDate').val("");
            } else if (retentionDate !== null && releaseDate >= retentionDate) {
                alert(releaseDateAlert_Retention);
                $('#ReleaseDate').val("");
            }
        }
    });
    //Retention Date validation where it should always be greater than or equal to today and always be greater than Release Date
    $('#RetentionDate').datepicker({
        onSelect: function () {
            var releaseDate = $('#ReleaseDate').datepicker('getDate');
            var retentionDate = $(this).datepicker('getDate');
            var todayDate = new Date();
            todayDate.setHours(0, 0, 0, 0);
            if (retentionDate < todayDate) {
                alert(retentionDateAlert_Today);
                $('#RetentionDate').val("");
            } else if (releaseDate !== null && releaseDate >= retentionDate) {
                alert(retentionDateAlert_Release);
                $('#RetentionDate').val("");
            }
        }
    });
}

// this function initiates the popup control 
function initiatePopUpControl() {
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', loadContentTypePopUp);
}

// this function format the RFG number in a particular format
function formatRFGNumber(number, length) {
    var str = '' + number;
    while (str.length < length) {
        str = '0' + str;
    }
    return 'rfg' + str;
}

/*
 *	function description: this function checks the file name and its extension and accordingly alert the user
 */
function checkFileUploadValidation() {
    // on change event for Native file control
    $("#NativeFileControlId").on('change', function () {
        let element = document.getElementById("NativeFileControlId");
        let existingAttachmentName_Native = $("#tblAttachmentNative").find("a").text();
        let index = existingAttachmentName_Native.lastIndexOf(".");
        let existingAttachmentName_NativeExtension = existingAttachmentName_Native.substring(index + 1, existingAttachmentName_Native.length);
        let file = element.files[0];
        let parts = element.value.split("\\");
        let fileName = parts[parts.length - 1];
        let lastIndexOfDot = fileName.lastIndexOf(".");
        let fileNameExtension = fileName.substring(lastIndexOfDot + 1, fileName.length);
        if (existingAttachmentName_Native !== "") {
            if (existingAttachmentName_Native !== fileName) {
                if (fileNameExtension.toLowerCase() === existingAttachmentName_NativeExtension.toLowerCase()) {
                    $("#txtNativeDocUploadType").addClass("showValidation").removeClass("hideValidation");
                    $("#txtNativeDocUploadType").text(fileNameMessage);
                } else {
                    $("#txtNativeDocUploadType").addClass("showValidation").removeClass("hideValidation");
                    $("#txtNativeDocUploadType").text(fileExtentionMessage);
                    $("#NativeFileControlId").val("");
                }
            } else {
                $("#txtNativeDocUploadType").addClass("hideValidation").removeClass("showValidation");
                $("#txtNativeDocUploadType").text("");
            }
        }
    });
    // on change event for Consumer file control
    $("#ConsumerFileControlId").on('change', function () {
        let element = document.getElementById("ConsumerFileControlId");
        let existingAttachmentName_Consumer = $("#tblAttachmentConsumer").find("a").text();
        let index = existingAttachmentName_Consumer.lastIndexOf(".");
        let existingAttachmentName_ConsumerExt = existingAttachmentName_Consumer.substring(index + 1, existingAttachmentName_Consumer.length);
        let file = element.files[0];
        let parts = element.value.split("\\");
        let fileName = parts[parts.length - 1];
        fileName = lastItemIdOfLib + "-" + fileName;
        let lastIndexOfDot = fileName.lastIndexOf(".");
        let fileNameExtension = fileName.substring(lastIndexOfDot + 1, fileName.length);
        if (existingAttachmentName_Consumer !== "") {
            if (existingAttachmentName_Consumer !== fileName) {
                if (fileNameExtension.toLowerCase() === existingAttachmentName_ConsumerExt.toLowerCase()) {
                    $("#txtConsumerDocUploadType").addClass("showValidation").removeClass("hideValidation");
                    $("#txtConsumerDocUploadType").text(fileNameMessage);
                } else {
                    $("#txtConsumerDocUploadType").addClass("showValidation").removeClass("hideValidation");
                    $("#txtConsumerDocUploadType").text(fileExtentionMessage);
                    $("#ConsumerFileControlId").val("");
                }
            } else {
                $("#txtConsumerDocUploadType").addClass("hideValidation").removeClass("showValidation");
                $("#txtConsumerDocUploadType").text("");
            }
        }
    });
}

/*
 *	function description: this function creates structure for Content Type popup
 */
function loadContentTypePopUp() {
    var htmlDialogModal = "<div id='itemRequestTypePopUp'>" +
        "<table  style='width:400px' cellpadding='4' align='center'>" +
        "<tr>" +
        "<td colspan='2' class='ms-formlabel' align='middle'><div>Select Content Type:</div></td>" +
        "<td colspan='4'><select id='itemRequestTypePopUp1'><option value='0'>Select</option></Select></td>" +
        "</tr>" +
        "</table>" +
        "</div>";
    $('body').append(htmlDialogModal);
    var mdOptions = {
        html: document.getElementById('itemRequestTypePopUp'),
        allowMaximize: false,
        showClose: true,
        width: 400,
        dialogReturnValueCallback: onClose
    };
    SP.UI.ModalDialog.showModalDialog(mdOptions);
    //calling this function to add content types on to the pop up
    getContentTypes();
    // on change of radion buttons
    $("[id='itemRequestTypePopUp1']").change(function () {
        var contentType = $("#itemRequestTypePopUp1 option:selected").text();
        $("#lblContentType").text(contentType);
        SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.Cancel);
        //call this function to hide/show fields on the basis of radio button selected
        changeFormFields(contentType);
    });
}

/*
 *	function description: this function redirects the user to the master library on closing the popup without selecting any content type
 */
function onClose(result, retValue) {
    if (result == SP.UI.DialogResult.cancel) {
        window.location.href = _spPageContextInfo.webAbsoluteUrl + "/" + masterLibraryName;
    }
}

/*
 *	function description: this function gets all the term sets and then retrieves terms of each term set
 */

function getTermSets() {
    var deferred = $.Deferred();
    var context = SP.ClientContext.get_current();
    var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
    //Name of the Term Store from which to get the Terms.  
    var termStore = session.getDefaultSiteCollectionTermStore();
    // var termStoreGroup = termStore.getSiteCollectionGroup(siteURL);
    //Name of the Term Group from which to get the Term sets.
    var termStoreGroup = termStore.getGroup("c36a38d6-e852-4e84-b38d-24a889d58bfb")
    //getting all the termsets of a particular group
    var parentTermSets = termStoreGroup.get_termSets();
    context.load(parentTermSets);
    context.executeQueryAsync(function onSuccess() {
        deferred.resolve(parentTermSets);
    }, function (sender, args) {
        deferred.reject(args.get_message());
    });
    return deferred.promise();
}

/*
 *	function description: this function is used to process termsets 
 */
function processTermSets() {
    if (formMode === "editForm") {
        popUpOpenRequestStarted();
    }
    var termSetIds = [];
    var asyncCallbacks = [];
    var termSetPromise = getTermSets();
    termSetPromise.done(function (result) { // Done means: When the promise has been fullfilled
        var termSetEnumerator = result.getEnumerator();
        while (termSetEnumerator.moveNext()) {
            var spTermSet = termSetEnumerator.get_current();
            var name = spTermSet.get_name();
            var id = spTermSet.get_id()._m_guidString$p$0;
            var terms = spTermSet.getAllTerms();
            asyncCallbacks.push(getTerms(terms, name));
        }
        // $.When is used to wait till all the callback calls are done and then only get the related terms
        $.when.apply($, asyncCallbacks).done(function () {
            console.log("all callbacks resolved");
            if (formMode === "editForm" || (formMode === "newForm" && itemId !== "")) {
                getMetadataFromList();
            }
        });
    });
}

/*
 *	function description: this function is used to get terms from indivisual termsets using async query with deferred promise
 */
function getTerms(terms, name) {
    var d = $.Deferred(),
        ddlDocumentType = document.getElementById("ddlDocumentType"),
        ddlBulletinType = document.getElementById("ddlBulletinType"),
        ddlBulletinType1 = document.getElementById("ddlBulletinType1"),
        ddlCommunicationType = document.getElementById("ddlCommunicationType"),
        ddlCommunicationType1 = document.getElementById("ddlCommunicationType1"),
        ddlCustomer = document.getElementById("ddlCustomer"),
        ddlCustomer1 = document.getElementById("ddlCustomer1"),
        ddlFirmware = document.getElementById("ddlFirmware"),
        ddlFirmware1 = document.getElementById("ddlFirmware1"),
        ddlLanguage = document.getElementById("ddlLanguage"),
        ddlManualType = document.getElementById("ddlManualType"),
        ddlManualType1 = document.getElementById("ddlManualType1"),
        ddlProductType = document.getElementById("ddlProductType"),
        ddlSecurityGroup = document.getElementById("ddlSecurityGroup"),
        context = SP.ClientContext.get_current();
    context.load(terms);
    context.executeQueryAsync(function () {
        var termsEnumerator = terms.getEnumerator();
        while (termsEnumerator.moveNext()) {
            var term = termsEnumerator.get_current();
            var option = document.createElement("OPTION");
            //optionDup is created as we have two different dropdowns of Bulletin, Comm etc. 
            var optionDup = document.createElement("OPTION");
            option.innerHTML = term.get_name();
            optionDup.innerHTML = term.get_name();
            option.value = term.get_id()._m_guidString$p$0;
            optionDup.value = term.get_id()._m_guidString$p$0;
            //Add the Option element to DropDownList.
            if (name === "Bulletin Type") {
                ddlBulletinType1.options.add(optionDup);
                ddlBulletinType.options.add(option);
            } else if (name === "Communication Type") {
                ddlCommunicationType.options.add(option);
                ddlCommunicationType1.options.add(optionDup);
            } else if (name === "Customer List") {
                ddlCustomer.options.add(option);
                ddlCustomer1.options.add(optionDup);
            } else if (name === "Document Type") {
                ddlDocumentType.options.add(option);
            } else if (name === "Firmware Type") {
                ddlFirmware.options.add(option);
                ddlFirmware1.options.add(optionDup);
            } else if (name === "Language") {
                ddlLanguage.options.add(option);
            } else if (name === "Manual Type") {
                ddlManualType.options.add(option);
                ddlManualType1.options.add(optionDup);
            } else if (name === "Product Type") {
                ddlProductType.options.add(option);
            } else if (name === "Security Group") {
                ddlSecurityGroup.options.add(option);
            }
        }
        d.resolve();
    });
    return d.promise();
}

/*
 *	function description: this function is used to get content types from content type configuration list
 */
function getContentTypes() {
    var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
    var itemsArray = [];
    var apiPath = siteURL + "/_api/web/lists/getbytitle('" + contentTypeListName + "')/items?$select=Title";
    $.ajax({
        url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
            if (data.d.results.length > 0) {
                var resultsColl = data.d.results;
                var ddlContentType = document.getElementById("itemRequestTypePopUp1");
                ddlContentType.append("<option>");
                for (var i = 0; i < resultsColl.length; i++) {
                    var option = document.createElement("OPTION");
                    option.innerHTML = resultsColl[i].Title;
                    option.value = resultsColl[i].Title;
                    ddlContentType.options.add(option);
                }
            }
        },
        error: function (data) {
            console.log("An error occurred. Please try again.");
        }
    });
}

/*
 *	function description: this function is used to get Offline types from Ricoh Offline type configuration list
 */
function getOfflineTypes() {
    var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
    var apiPath = siteURL + "/_api/web/lists/getbytitle('" + offlineTypeListName + "')/items?$select=Title,Id";
    $.ajax({
        url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
            if (data.d.results.length > 0) {
                var resultsColl = data.d.results;
                var ddlOfflineType = document.getElementById("ddlOfflineType");
                for (var i = 0; i < resultsColl.length; i++) {
                    var option = document.createElement("OPTION");
                    option.innerHTML = resultsColl[i].Title;
                    option.value = resultsColl[i].Title;
                    ddlOfflineType.options.add(option);
                }
            }
        },
        eror: function (data) {
            console.log("An error occurred. Please try again.");
        }
    });
}

/*
 *	function description: this function is used to disable/enable fields on basis of dropdown selection
 */
function changeFormFields(CTValue) {
    if (CTValue === "Bulletins") { //for Bulletin Type
        $(".ddlBulletinType, .txtComments, .txtKeywords").show();
        $("#ddlBulletinType1, #ddlCustomer1, #ddlFirmware1, #ddlManualType1, #ddlCommunicationType1 option:selected").val(0);
        $(".ddlCustomer, .ddlFirmware, .ddlCommunicationType, .ddlBulletinType1, .ddlCustomer1, .ddlFirmware1, .ddlManualType,.ddlManualType1, .ddlCommunicationType1").hide();
    } else if (CTValue === "Standard") { //for Standard Type
        $(".ddlBulletinType1, .ddlCustomer1, .ddlFirmware1, .ddlManualType1, .ddlCommunicationType1, .txtComments, .txtKeywords").show();
        $(".ddlBulletinType, .ddlCustomer, .ddlFirmware, .ddlManualType, .ddlCommunicationType").hide();
        $("#ddlBulletinType, #ddlCustomer, #ddlFirmware, #ddlManualType, #ddlCommunicationType option:selected").val(0);
    } else if (CTValue === "Communication") { //for Communication Type
        $(".ddlCommunicationType,.txtComments, .txtKeywords").show();
        $(".ddlBulletinType1,.ddlBulletinType, .ddlCustomer1, .ddlFirmware1, .ddlManualType1, .ddlCommunicationType1").hide();
        $("#ddlBulletinType1, #ddlCustomer1, #ddlFirmware1, #ddlManualType1, #ddlCommunicationType1, #ddlBulletinType option:selected").val(0);
    } else if (CTValue === "Firmware") { //for Firmware Type
        $(".ddlFirmware, .ddlCustomer1,.txtComments, .txtKeywords").show();
        $(".ddlBulletinType, .ddlCustomer, .ddlManualType, .ddlCommunicationType").hide();
        $("#ddlBulletinType, #ddlCustomer, #ddlManualType, #ddlCommunicationType option:selected").val(0);
        $(".ddlBulletinType1, .ddlFirmware1, .ddlManualType1, .ddlCommunicationType1").hide();
        $("#ddlBulletinType1, #ddlFirmware1, #ddlManualType1, #ddlCommunicationType1 option:selected").val(0);

    } else if (CTValue === "Service Docs") { //for Service Docs Type
        $(".ddlManualType, .ddlCommunicationType1,.ddlCustomer1").show();
        $(".ddlBulletinType, .ddlCustomer, .ddlFirmware, .ddlCommunicationType,.ddlBulletinType1, .ddlFirmware1, .ddlManualType1").hide();
        $("#ddlBulletinType, #ddlCustomer, #ddlFirmware, #ddlCommunicationType,#ddlBulletinType1, #ddlFirmware1, #ddlManualType1 option:selected").val(0);
        $(".txtComments, .txtKeywords").hide();
    } else if (CTValue === "Custom Config") { //for Custom Config Type
        $(".ddlBulletinType1, .ddlCustomer, .ddlFirmware1, .ddlManualType1, .ddlCommunicationType1, .txtComments, .txtKeywords").show();
        $(".ddlBulletinType, .ddlFirmware, .ddlCustomer1, .ddlManualType, .ddlCommunicationType").hide();
        $("#ddlBulletinType, #ddlCustomer1, #ddlFirmware, #ddlManualType, #ddlCommunicationType option:selected").val(0);
    }
}

/*
 *	function description: this function gets the last sharepoint item ID from the master library which is used in creating unique RFG No
 */
function getLastItemIdFromLib() {
    var apiPath = siteURL + "/_api/web/lists/getbytitle('" + masterLibraryName + "')/items?$select=ID&$orderBy=ID desc&$top=1";
    $.ajax({
        url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
            if (data.d.results.length > 0) {
                var resultsColl = data.d.results[0];
                var LastListItemID = resultsColl.ID + 1;
                lastItemIdOfLib = formatRFGNumber(LastListItemID, 6);
            }
        },
        eror: function (data) {
            console.log("An error occurred. Please try again.");
        }
    });
}

/*
 *	function description: this function checks if the folder exists or not. if not, it creates the folder in library
 */
function checkFolderExists(folderNameByProduct, fileUploadControlID) {
    var ctx = SP.ClientContext.get_current();
    var folder = ctx.get_web().getFolderByServerRelativeUrl(siteServerURL + "/" + masterLibraryName + "/" + folderNameByProduct);
    ctx.load(folder, "Exists", "Name");
    ctx.executeQueryAsync(
        Function.createDelegate(this, onSuccess),
        Function.createDelegate(this, onFail)
    );

    function onSuccess() {
        if (folder.get_exists()) {
            addContentType(fileUploadControlID);
        }
        else {
            createFolder(folderNameByProduct, fileUploadControlID);
        }
    }

    function onFail(s, args) {
        if (args.get_errorTypeName() === "System.IO.FileNotFoundException") {
            console.log("Folder does not exist.");
            createFolder(folderNameByProduct, fileUploadControlID);
        }
        else {
            console.log("Error: " + args.get_message());
        }
    }
}

/*
 *	function description: this function creates folder in the doc lib
 */
function createFolder(folderNameByProduct, fileUploadControlID) {
    var clientContext = new SP.ClientContext.get_current();
    var oWebsite = clientContext.get_web();
    var oList = oWebsite.get_lists().getByTitle(masterLibraryName);
    var itemCreateInfo = new SP.ListItemCreationInformation();
    itemCreateInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
    itemCreateInfo.set_leafName(folderNameByProduct);
    this.oListItem = oList.addItem(itemCreateInfo);
    this.oListItem.update();

    clientContext.load(this.oListItem);
    clientContext.executeQueryAsync(
        Function.createDelegate(this, successHandler),
        Function.createDelegate(this, errorHandler)
    );

    function successHandler() {
        console.log("Folder Created Successfully!");
        // on creation of folder, add content type to the library
        addContentType(fileUploadControlID);
    }

    function errorHandler() {
        console.log("Request failed: " + arguments[1].get_message());
        popUpRequestEnded();
    }
}
 
/*
 *	function description: This method is used to add the content type on exisiting list. 
 */
function addContentType(fileUploadControlID) {
    //Get host web URL  
    var context = SP.ClientContext.get_current();
    var oList = context.get_web().get_lists().getByTitle(masterLibraryName);
    var allContentTypeColl = oList.get_contentTypes();
    //Get all content types from list level  
    context.load(allContentTypeColl);
    context.load(oList);
    context.executeQueryAsync(function () {
        var CTypeID;
        var contentTypeName = $("#lblContentType").text();
        var contentTypeEnum = allContentTypeColl.getEnumerator();
        while (contentTypeEnum.moveNext()) {
            var currentCT = contentTypeEnum.get_current();
            if (currentCT.get_name() == contentTypeName) {
                CTypeID = currentCT.get_stringId();
                break;
            }
        }
        prepareMasterData(fileUploadControlID, CTypeID);
    }, function (sender, args) {
        onfail(sender, args);
    });
}

// This function is executed if the above call fails  
function onfail(sender, args) {
    alert('Failed to add content type in list. Error:' + args.get_message());
}

/*
 *	function description: this function is used to validate the data before saving it into the master list
 */
function checkValidation() {
    var validationFlag = true,
        selectedContentTypeValue = $("#lblContentType").text(),
        title = $("#Title").val(),
        documentType = $("#ddlDocumentType option:selected").text(),
        bulletinType = $("#ddlBulletinType option:selected").text(),
        communicationType = $("#ddlCommunicationType option:selected").text(),
        manualType = $("#ddlManualType option:selected").text(),
        firmware = $("#ddlFirmware option:selected").text(),
        customer = $("#ddlCustomer option:selected").text(),
        nativeDocumentUploadPath = $("#NativeFileControlId").val();
    existingAttachment = $('#tblAttachmentNative tr td').length;
    //checking validation for controls which are common for all the selected Content Types
    if (selectedContentTypeValue === "Standard" || selectedContentTypeValue === "Bulletins" || selectedContentTypeValue === "Communication" || selectedContentTypeValue === "Custom Config" || selectedContentTypeValue === "Firmware" || selectedContentTypeValue === "Service Docs") {
        if (title === "") {
            validationFlag = false;
            $("#txtTitle").addClass("showValidation").removeClass("hideValidation");
            $("#txtTitle").text("This field is mandatory.");
        } else {
            $("#txtTitle").addClass("hideValidation").removeClass("showValidation");
            $("#txtTitle").text("");
        }
        if (documentType === "Select") {
            validationFlag = false;
            $("#txtDocType").addClass("showValidation").removeClass("hideValidation");
            $("#txtDocType").text("This field is mandatory.");
        } else {
            $("#txtDocType").addClass("hideValidation").removeClass("showValidation");
            $("#txtDocType").text("");
        }
        if (nativeDocumentUploadPath === "" && existingAttachment < 1) {
            validationFlag = false;
            $("#txtNativeDocUploadType").addClass("showValidation").removeClass("hideValidation");
            $("#txtNativeDocUploadType").text("This field is mandatory.");
        } else {
            $("#txtNativeDocUploadType").addClass("hideValidation").removeClass("showValidation");
            $("#txtNativeDocUploadType").text("");
        }
    }
    //checking validation for Bulletin Type
    if (selectedContentTypeValue === "Bulletins") {
        if (bulletinType === "Select") {
            validationFlag = false;
            $("#txtBulletinType").addClass("showValidation").removeClass("hideValidation");
            $("#txtBulletinType").text("This field is mandatory.");
        } else {
            $("#txtBulletinType").addClass("hideValidation").removeClass("showValidation");
            $("#txtBulletinType").text("");
        }
    }
    //checking validation for Communication Type
    if (selectedContentTypeValue === "Communication") {
        if (communicationType === "Select") {
            validationFlag = false;
            $("#txtCommType").addClass("showValidation").removeClass("hideValidation");
            $("#txtCommType").text("This field is mandatory.");
        } else {
            $("#txtCommType").addClass("hideValidation").removeClass("showValidation");
            $("#txtCommType").text("");
        }
    }
    //checking validation for Custom Config
    if (selectedContentTypeValue === "Custom Config") {
        if (customer === "Select") {
            validationFlag = false;
            $("#txtCustomerType").addClass("showValidation").removeClass("hideValidation");
            $("#txtCustomerType").text("This field is mandatory.");
        } else {
            $("#txtCustomerType").addClass("hideValidation").removeClass("showValidation");
            $("#txtCustomerType").text("");
        }
    }
    //checking validation for Firmware
    if (selectedContentTypeValue === "Firmware") {
        if (firmware === "Select") {
            validationFlag = false;
            $("#txtFirmware").addClass("showValidation").removeClass("hideValidation");
            $("#txtFirmware").text("This field is mandatory.");
        } else {
            $("#txtFirmware").addClass("hideValidation").removeClass("showValidation");
            $("#txtFirmware").text("");
        }
    }
    //checking validation for Service Docs
    if (selectedContentTypeValue === "Service Docs") {
        if (manualType === "Select") {
            validationFlag = false;
            $("#txtManualType").addClass("showValidation").removeClass("hideValidation");
            $("#txtManualType").text("This field is mandatory.");
        } else {
            $("#txtManualType").addClass("hideValidation").removeClass("showValidation");
            $("#txtManualType").text("");
        }
    }
    return validationFlag;
}

/*
 *	function description: this function is used to save Master Data on submit button click
 */
function saveMasterData(fileUploadControlID) {
    popUpOpenRequestStarted();
    var selectedProductCode = $("#ddlProductType option:selected").text();
    if (selectedProductCode !== "Select") {
        checkFolderExists(selectedProductCode, fileUploadControlID);
    } else {
        addContentType(fileUploadControlID);
    }
}

/*
 *	function description: this function closes the wait screen popup
 */
function popUpRequestEnded(sender, args) {
    try {
        waitDialog.close();
        waitDialog = null;
    } catch (ex) { }
};

/*
 *	function description: this function open the wait screen popup
 */
function popUpOpenRequestStarted(sender, args) {
    ExecuteOrDelayUntilScriptLoaded(ShowWaitDialog, "sp.js");
};

function ShowWaitDialog() {
    try {
        if (waitDialog == null) {
            waitDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose('Processing...', 'Please wait while request is in progress...', 76, 330);
        }
    } catch (ex) { }
};

function prepareMasterData(fileUploadControlID, contentTypeID) {
    //check if validation of different controls is passed or not
    if (checkValidation()) {
        var formattedOfflineType = "", offlineType;
        offlineType = $("#ddlOfflineType").val() == null ? "" : $("#ddlOfflineType").val();
        if (offlineType !== "") {
            for (let counter = 0; counter < offlineType.length; counter += 1) {
                formattedOfflineType += offlineType[counter] + ",";
            }
            formattedOfflineType = formattedOfflineType.replace(/,\s*$/, "");
        }
        var itemProps = {
            'Title': $('#Title').val(),
            'OffLine_x0020_Type': formattedOfflineType,
            'RicohComments': $('#txtComments').val() || '',
            'Engineering': $('#txtEngineering').val() || '',
            'Model_x0020_Name': $('#txtModalName').val() || '',
            'Product_x0020_Code': $('#txtProductCode').val() || '',
            'Keywords': $('#txtKeywords').val() || '',
            'OptionCode': $('#txtOptionCode').val() || '',
            'Release_x0020_Date': $("#ReleaseDate").val() === "" ? null : new Date($("#ReleaseDate").val()).toLocaleDateString(),
            'Retention_x0020_Period': $("#RetentionDate").val() === "" ? null : $("#RetentionDate").val(),
            'ContentTypeId': contentTypeID,
            'RFGSrNo': lastItemIdOfLib == undefined ? $("#lblRFGSerialNo").text() : lastItemIdOfLib
        };

        if (fileUploadControlID == "NativeFileControlId" ) {
        	if (formMode === "newForm") {
        		itemProps["isNative"] = "1";
            	itemProps["Legacy_Info"] = "B";
        	}else {
        		itemProps["isNative"] = isNative;
        		itemProps["ConsumerDocumentID"] = consumerDocumentId;
        	}
            setDocumentStatus(itemProps);
        } else if (fileUploadControlID == "ConsumerFileControlId" && (nativeDocumentId == "") && consumerDocumentId == null) {
            itemProps["IsConsumer"] = "1";
            itemProps["NativeDocumentID"] = itemId;
            itemProps["Legacy_Info"] = "B";
            itemProps["DocumentStatus"] = "Released";
        }
        readFile(itemProps, fileUploadControlID);
    } else {
        //close the message pop up 
        popUpRequestEnded();
    }
}

/*
 *	function description: this function is used to read attached files
 */
function readFile(itemProperties, fileUploadControlID) {
    var itemProps = itemProperties;
    //Get File Input Control and read the file name  
    var element = document.getElementById(fileUploadControlID);
    if (fileUploadControlID === "NativeFileControlId") {
        existingFileID = "#tblAttachmentNative";
    } else {
        existingFileID = "#tblAttachmentConsumer";
    }
    var existingAttachmentNative = $(existingFileID).find("a").text();
    var file = element.files[0];
    var parts = element.value.split("\\");
    var fileName = parts[parts.length - 1];
    if (formMode === "editForm") {
        if (fileUploadControlID === "NativeFileControlId") {
            if (fileName !== "" && $("#controlsDiv").data("changed") == true) {
                itemProps["Legacy_Info"] = "B"; // B means Both the file and Metadata have changed in edit form              
            } else if (fileName === "" && $("#controlsDiv").data("changed") == true) {
                itemProps["Legacy_Info"] = "M";   // M means only Metadata have changed in edit form         
            } else if (fileName !== "" && ($("#controlsDiv").data("changed") == undefined || $("#controlsDiv").data("changed") == false)) {
                itemProps["Legacy_Info"] = "F";   // F means only the file have changed in edit form
            } else {
                itemProps["Legacy_Info"] = legacyInfo;    // "" means nothing have changed
            }
            setDocumentStatus(itemProps);
        } else {
            if (fileName !== "" && $("#controlsDiv").data("changed") == true) {
                itemProps["Legacy_Info"] = "B";
            } else if (fileName === "" && $("#controlsDiv").data("changed") == true) {
                itemProps["Legacy_Info"] = "M";
            } else if (fileName !== "" && ($("#controlsDiv").data("changed") == false || $("#controlsDiv").data("changed") == undefined)) {
                itemProps["Legacy_Info"] = "F";
            } else {
                itemProps["Legacy_Info"] = legacyInfo;
            }
            // for consumer type file, directly change the doocumentStatus to Released
            itemProps["DocumentStatus"] = "Released";
        }
    }
    //check if both the files are present for a single control
    if (existingAttachmentNative !== "" && fileName !== "") {
        if (fileUploadControlID === "NativeFileControlId") {
            if (existingAttachmentNative !== fileName) {
                fileName = existingAttachmentNative;
            }
        } else if (existingAttachmentNative.split("-")[1] !== fileName){ 
	            fileName = existingAttachmentNative;
	    } else if (existingAttachmentNative.split("-")[1] === fileName) {
	        	fileName = existingAttachmentNative;
	        }
    }
    //add prefix as "rfg number" for consumer related file names
    if (fileUploadControlID === "ConsumerFileControlId" && file !== undefined && existingAttachmentNative === "") {
        fileName = lastItemIdOfLib + "-" + fileName;
    }
    //directly call saveMetadataToList function when there is no file to upload
    if (file == undefined) {
        saveMetadataToList(itemId, itemProps, fileUploadControlID);
    } else {
        uploadDocument(file, fileName, itemProps, fileUploadControlID);
    }
}
/*
 *	function description: this function set the document status value on the basis of different conditions
 */
function setDocumentStatus(itemProps) {
    var docType = $("#ddlDocumentType option:selected").text() == "Select" ? "" : $("#ddlDocumentType option:selected").text();
    var CommType = $("#ddlCommunicationType option:selected").text() == "Select" ? "" : $("#ddlCommunicationType option:selected").text();
    if (itemProps.Legacy_Info === "B" || itemProps.Legacy_Info === "F") {
        if (docType === "Bulletins" && CommType !== "Training Announcement") {
            checkCurentUserInSPGroup(SPKMGroup).then(function (data) {
                if (data.d.results.length > 0) {
                    itemProps["DocumentStatus"] = "Pending";
                } else {
                    itemProps["DocumentStatus"] = "Released";
                }
            });
        } else if (CommType === "Training Announcement" && docType !== "Bulletins") {
            checkCurentUserInSPGroup(SPCDGroup).then(function (data) {
                if (data.d.results.length > 0) {
                    itemProps["DocumentStatus"] = "Pending";
                } else {
                    itemProps["DocumentStatus"] = "Released";
                }
            });
        } else {
            itemProps["DocumentStatus"] = "Released";
        }
    } else {
        itemProps["DocumentStatus"] = "Released";
    }
}

var myListItem;
function saveMetadataToList(itemId, itemProps, fileUploadControlID) {
    //Get Client Context,Web and List object.  
    var clientContext = SP.ClientContext.get_current();
    var oList = clientContext.get_web().get_lists().getByTitle(masterLibraryName);
    var contentTypes = oList.get_contentTypes();
    var field = oList.get_fields().getByInternalNameOrTitle("Bulletin_x0020_Type");
    var field1 = oList.get_fields().getByInternalNameOrTitle("Document_x0020_Type");
    var field2 = oList.get_fields().getByInternalNameOrTitle("Communication_x0020_Type");
    var field3 = oList.get_fields().getByInternalNameOrTitle("Customer");
    var field4 = oList.get_fields().getByInternalNameOrTitle("Firmware");
    var field5 = oList.get_fields().getByInternalNameOrTitle("Language_x0020_Type");
    var field6 = oList.get_fields().getByInternalNameOrTitle("Manual_x0020_Type");
    var field7 = oList.get_fields().getByInternalNameOrTitle("Product_x0020_Type");
    var field8 = oList.get_fields().getByInternalNameOrTitle("Security_x0020_Group");
    var txField = clientContext.castTo(field, SP.Taxonomy.TaxonomyField);
    var txField1 = clientContext.castTo(field1, SP.Taxonomy.TaxonomyField);
    var txField2 = clientContext.castTo(field2, SP.Taxonomy.TaxonomyField);
    var txField3 = clientContext.castTo(field3, SP.Taxonomy.TaxonomyField);
    var txField4 = clientContext.castTo(field4, SP.Taxonomy.TaxonomyField);
    var txField5 = clientContext.castTo(field5, SP.Taxonomy.TaxonomyField);
    var txField6 = clientContext.castTo(field6, SP.Taxonomy.TaxonomyField);
    var txField7 = clientContext.castTo(field7, SP.Taxonomy.TaxonomyField);
    var txField8 = clientContext.castTo(field8, SP.Taxonomy.TaxonomyField);
    myListItem = oList.getItemById(itemId);

    // set all the form fields in for loop
    for (var propName in itemProps) {
        myListItem.set_item(propName, itemProps[propName]);
    }
    var termFieldValue = new SP.Taxonomy.TaxonomyFieldValue();
    //loop thorugh all the dropdown fields with class taxField(this class is only assigned to taxonomy related controls)
    $('select.taxField').each(function () {
        var taxFieldGuid = $(this).val();
        var taxFieldLabel = $(this).find('option:selected').text();
        //check if selected dropdown value is not blank or select or 0
        if (taxFieldLabel !== "Select" && taxFieldLabel !== "" && taxFieldGuid !== "0") {
            var selectedDropdownId = $(this).attr('id');
            termFieldValue.set_label(taxFieldLabel);
            termFieldValue.set_termGuid(taxFieldGuid);
            termFieldValue.set_wssId(-1);
            if (selectedDropdownId === "ddlBulletinType" || selectedDropdownId === "ddlBulletinType1") {
                txField.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlDocumentType") {
                txField1.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlCommunicationType" || selectedDropdownId === "ddlCommunicationType1") {
                txField2.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlCustomer" || selectedDropdownId === "ddlCustomer1") {
                txField3.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlFirmware" || selectedDropdownId === "ddlFirmware1") {
                txField4.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlLanguage") {
                txField5.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlProductType") {
                txField7.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlSecurityGroup") {
                txField8.setFieldValueByValue(myListItem, termFieldValue);
            } else if (selectedDropdownId === "ddlManualType" || selectedDropdownId === "ddlManualType1") {
                txField6.setFieldValueByValue(myListItem, termFieldValue);
            }
        }
    });
    //callingSubmitFunctionFlag is introduced to check if this function does not loop infinitely. 
    if (fileUploadControlID == "NativeFileControlId") {
        callingSubmitFunctionFlag = 1; //for Native File, its value is set as 1 
    } else {
        callingSubmitFunctionFlag = 2;  //for Consumer File, its value is set as 2
    }
    myListItem.update();
    clientContext.load(myListItem);
    clientContext.executeQueryAsync(Function.createDelegate(this, this.QuerySuccess), Function.createDelegate(this, this.QueryFailure));
}

function QuerySuccess(sender, args) {
    itemId = myListItem.get_id();
    var isConsumer = myListItem.get_fieldValues().IsConsumer,
        isNative = myListItem.get_fieldValues().isNative,
        nativeDocId = myListItem.get_fieldValues().NativeDocumentID,
        consumerDocId = myListItem.get_fieldValues().ConsumerDocumentID,
        existingConsumerDoc = $("#tblAttachmentConsumer").find("a").text(),
        newProductCode = $("#ddlProductType option:selected").text();
    //setting global variables
    nativeDocumentId = nativeDocId || "";
    consumerDocumentId = consumerDocId;
    if (isConsumer == true) {
        itemId = nativeDocId;
        var newItemId = myListItem.get_id(); // this is specially created for UpdateConsumerID function
        if (consumerDocId === null && callingSubmitFunctionFlag == 2) {
            checkConsumerIDExists(newItemId, nativeDocId);
        }
    } else if (isNative == true) {
        itemId = consumerDocId == null ? myListItem.get_id() : consumerDocId;
    }
    //recursively call this function if there is Consumer file to upload as well
    if (callingSubmitFunctionFlag == 1 && ($("#ConsumerFileControlId").val() !== ""  || (existingConsumerDoc !== ""  && $("#controlsDiv").data("changed") == true )) ){
        saveMasterData("ConsumerFileControlId");
    } else {
        // existingProductType flag is undefined when form is in newForm Mode.
        if (existingProductType !== undefined) {
            if (existingProductType !== newProductCode && ((callingSubmitFunctionFlag == 2 && formMode !== "newForm") || ($("#ConsumerFileControlId").val() !== "" && existingConsumerDoc !== "") || (callingSubmitFunctionFlag == 1 && $("#ConsumerFileControlId").val() == "" && existingConsumerDoc == ""))) {
                moveFiles();
            } else if (existingProductType === newProductCode && (callingSubmitFunctionFlag == 2 || callingSubmitFunctionFlag == 1)) {
                popUpRequestEnded();
                setTimeout(function () {
                    alert("Document is saved successfully");
                    redirect();
                }, 300);
            }
        } else {
            popUpRequestEnded();
            setTimeout(function () {
                alert("Document is saved successfully");
                redirect();
            }, 300);
        }
    }
}

function QueryFailure(sender, args) {
    console.log('Request failed with error message - ' + args.get_message());
    popUpRequestEnded();
    //alert("Request failed with error message - " + args.get_message());
    alert("Something went wrong, please refresh the page and try again.");
}

var files;
function moveFiles() {
    var context = new SP.ClientContext.get_current(),
        newFolder = $("#ddlProductType option:selected").text(),
        fileNameNative = $("#tblAttachmentNative").find("a").text(),
        fileNameConsumer = $("#tblAttachmentConsumer").find("a").text(),
        web = context.get_web();
    existingProductType = existingProductType == "Select" ? "" : existingProductType;
    var folder = web.getFolderByServerRelativeUrl(siteServerURL + "/" + masterLibraryName + "/" + existingProductType);
    files = folder.get_files();
    context.load(files);
    context.executeQueryAsync(
        function () {
            var e = files.getEnumerator();
            while (e.moveNext()) {
                var file = e.get_current();
                if (file.get_name() === fileNameNative || file.get_name() === fileNameConsumer) {
                    var destLibUrl = siteServerURL + "/" + masterLibraryName + "/" + newFolder + "/" + file.get_name();
                    file.moveTo(destLibUrl, SP.MoveOperations.overwrite);
                }
            }
            context.executeQueryAsync(function () {
                popUpRequestEnded();
                alert("File(s) moved successfully to folder " + newFolder);
                redirect();
            },
                function (sender, args) {
                    console.log("error: " + args.get_message());
                });
        },
        function (sender, args) {
            console.log("Sorry, something messed up: " + args.get_message());
            popUpRequestEnded();
        }
    );
}
function checkConsumerIDExists(itemId, nativeDocId) {
    var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
    var apiPath = siteURL + "/_api/web/lists/getbytitle('" + masterLibraryName + "')/items(" + nativeDocId + ")?$select=ConsumerDocumentID";
    $.ajax({
        url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
            let consumerId = data.d.ConsumerDocumentID;
            //check if Consumer ID is not null or empty or undefined or false
            if (consumerId === null) {
                // update Consumer Document ID to the related native Document
                UpdateConsumerID(itemId, nativeDocId);
            }
        },
        eror: function (data) {
            console.log("An error occurred. Please try again.");
        }
    });
}

function UpdateConsumerID(itemId, nativeDocId) {
    siteURL = _spPageContextInfo.webAbsoluteUrl;
    var apiPath = siteURL + "/_api/web/lists/getbytitle('" + masterLibraryName + "')/items/getbyid(" + nativeDocId + ")";
    $.ajax({
        url: apiPath,
        type: "POST",
        headers:
        {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        },
        data: JSON.stringify
            ({
                __metadata:
                {
                    type: "SP.Data.ESource_x0020_Content_x0020_LibraryItem"
                },
                "ConsumerDocumentID": itemId
            }),
        async: false,
        success: function (data) {
            console.log("Item updated successfully");
        }, eror: function (data) {
            console.log("An error occurred. Please try again.");
        }
    })
}

function getMetadataFromList() {
    var tblRows = "";
    var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
    var apiPath = siteURL + "/_api/web/lists/getByTitle('" + masterLibraryName + "')/Items(" + itemId + ")?$select=*,Author/Title,Editor/Title,FileLeafRef,File/ServerRelativeUrl,ContentType/Id,ContentType/Name&$expand=File,Author,Editor,ContentType";
    $.ajax({
        url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
            if (data.d) {
                var resultsColl = data.d;
                //load taxonomy values
                getTaxonomyValue(resultsColl);
                //load values other than taxonomy values
                var releaseDate = stringToDate(isNonNull(resultsColl.Release_x0020_Date));
                var retentionDate = stringToDate(isNonNull(resultsColl.Retention_x0020_Period));
                var modifiedDate = stringToDate(isNonNull(resultsColl.Modified));
                var createdDate = stringToDate(isNonNull(resultsColl.Created));
                //setting up global variables
                existingProductType = $("#ddlProductType option:selected").text();
                isConsumer = resultsColl.IsConsumer;
                isNative = resultsColl.isNative;
                nativeDocumentId = resultsColl.NativeDocumentID;
                consumerDocumentId = resultsColl.ConsumerDocumentID;
                legacyInfo = resultsColl.Legacy_Info;
                if (formMode !== "newForm" && itemId !== "") {
                    lastItemIdOfLib = resultsColl.RFGSrNo;
                }
                if (resultsColl.OffLine_x0020_Type !== null) {
                    var offlineType = resultsColl.OffLine_x0020_Type.results[0];
                    var arrOfflineType = offlineType.split(',');
                    $("#ddlOfflineType").val(arrOfflineType);
                    $(".selectpicker").selectpicker("refresh");
                }
                $("#Title").val(resultsColl.Title);
                $("#lblCreatedBy").text(resultsColl.Author.Title);
                $("#lblModifiedBy").text(resultsColl.Editor.Title);
                $("#lblModifiedDate").text(modifiedDate);
                $("#lblCreatedDate").text(createdDate);
                $("#lblRFGSerialNo").text(resultsColl.RFGSrNo);
                $("#lblContentType").text(resultsColl.ContentType.Name);
                $("#ddlOfflineType").val(resultsColl.OffLine_x0020_Type === null ? 0 : resultsColl.OffLine_x0020_Type);
                $("#txtEngineering").val(resultsColl.Engineering);
                $("#txtModalName").val(resultsColl.Model_x0020_Name);
                $("#txtKeywords").val(resultsColl.Keywords);
                $("#txtOptionCode").val(resultsColl.OptionCode);
                $("#txtProductCode").val(resultsColl.Product_x0020_Code);
                $("#txtComments").val(resultsColl.RicohComments);
                $("#ReleaseDate").val(releaseDate);
                $("#RetentionDate").val(retentionDate);
                var selectedContentTypeValue = resultsColl.ContentType.Name;
                changeFormFields(selectedContentTypeValue);
                var filterUrl = "";
                //load attachements here
                if (formMode == "editForm") {
                    nativeFileFullPath = resultsColl.File.ServerRelativeUrl;
                    tblRows = "<tr><td><a href='" + resultsColl.File.ServerRelativeUrl + "'>" + resultsColl.FileLeafRef + "</td></tr>";
                    $(".basicInfo").show();
                    if (resultsColl.isNative === true) {
                        $("#tblAttachmentNative").append(tblRows);
                        filterUrl = "NativeDocumentID eq " + resultsColl.ID;
                        popUpRequestEnded();
                        if (resultsColl.DocumentStatus.toLowerCase() !== "released" && resultsColl.DocumentStatus.toLowerCase() !== "rejected") {
                            setTimeout(function () {
                                alert("This document is currently undergoing apporval process. Please edit the document once it is approved.");
                            }, 300);
                            $("#controlsDiv :input, #btnSubmit").prop("disabled", true);
                        }
                        getAttachmentAndDocStatus(filterUrl, "#tblAttachmentConsumer", "");
                    } else if (resultsColl.IsConsumer === true) {
                        $("#tblAttachmentConsumer").append(tblRows);
                        filterUrl = "ID eq " + resultsColl.NativeDocumentID;
                        popUpRequestEnded();
                        setTimeout(function () {
                            getAttachmentAndDocStatus(filterUrl, "#tblAttachmentNative", "disableFields");
                        }, 300);
                    } else {
                        popUpRequestEnded();
                    }
                }
            }
        },
        eror: function (data) {
            console.log("An error occurred. Please try again.");
            popUpRequestEnded();
        }
    });
}

function getAttachmentAndDocStatus(filterUrl, IdIdentifier, disableFieldsIdentifier) {
    var tblRows = "";
    var apiPath = siteURL + "/_api/web/lists/getByTitle('" + masterLibraryName + "')/Items?$select=ID,DocumentStatus,isNative,ConsumerDocumentID,FileLeafRef,File/ServerRelativeUrl&$expand=File&$filter=" + filterUrl;
    $.ajax({
        url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false,
        success: function (data) {
            if (data.d.results.length > 0) {
                var resultsColl = data.d.results[0];
                var docStatus = resultsColl.DocumentStatus;
                if (disableFieldsIdentifier === "disableFields") {
                	consumerDocumentId = resultsColl.ConsumerDocumentID;
                	isNative = resultsColl.isNative;
                	//this Id is set so that when Consumer doc is opened to edit, we will first update native doc and then Consumer Id 
                	itemId = resultsColl.ID;
                }
                consumerFileFullPath = resultsColl.File.ServerRelativeUrl;
                tblRows = "<tr><td><a href='" + resultsColl.File.ServerRelativeUrl + "'>" + resultsColl.FileLeafRef + "</td></tr>";
                $(IdIdentifier).append(tblRows);
                if (docStatus !== "Released" && disableFieldsIdentifier === "disableFields") {
                    alert("The Native document of this consumer file is currently undergoing apporval process. Please edit the document once it is approved.");
                    $("#controlsDiv :input, #btnSubmit").prop("disabled", true);
                }
            }
        },
        eror: function (data) {
            console.log("An error occurred. Please try again.");
        }
    });
}

function getTaxonomyValue(obj) {
    var metaString = "";
    // Iterate over the fields in the row of data
    for (var field in obj) {
        // If it's the field we're interested in....
        if (obj.hasOwnProperty(field)) {
            if (obj[field] !== null) {
                if (obj[field].WssId !== undefined) {
                    // get the WssId from the field ...
                    var thisId = obj[field].WssId;
                    var termGuid = obj[field].TermGuid;

                    switch (field) {
                        case "Bulletin_x0020_Type":
                            $("#ddlBulletinType").val(termGuid);
                            $("#ddlBulletinType1").val(termGuid);
                            break;
                        case "Document_x0020_Type":
                            $("#ddlDocumentType").val(termGuid);
                            break;
                        case "Communication_x0020_Type":
                            $("#ddlCommunicationType").val(termGuid);
                            $("#ddlCommunicationType1").val(termGuid);
                            break;
                        case "Customer":
                            $("#ddlCustomer").val(termGuid);
                            $("#ddlCustomer1").val(termGuid);
                            break;
                        case "Firmware":
                            $("#ddlFirmware").val(termGuid);
                            $("#ddlFirmware1").val(termGuid);
                            break;
                        case "Language_x0020_Type":
                            $("#ddlLanguage").val(termGuid);
                            break;
                        case "Manual_x0020_Type":
                            $("#ddlManualType").val(termGuid);
                            $("#ddlManualType1").val(termGuid);
                            break;
                        case "Product_x0020_Type":
                            $("#ddlProductType").val(termGuid);
                            break;
                        case "Security_x0020_Group":
                            $("#ddlSecurityGroup").val(termGuid);
                            break;
                    }
                }
            }
        }
    }
}


/*
 *	function description: this function uploads the attached document to the library
 */
function uploadDocument(file, fileName, itemProperties, fileUploadControlID) {
    var fileUploadService = new FileUploadService();
    fileUploadService.fileUpload(file, masterLibraryName, fileName).then(addFileToFolder => {
        console.log("File Uploaded Successfully");
        var parsedData = JSON.parse(addFileToFolder.body);
        var listItemAllFieldsURL = parsedData.d.ListItemAllFields.__deferred.uri;
        var getItem = getListItem(listItemAllFieldsURL);
        getItem.done(function (listItem, status, xhr) {
            var savedItemId = listItem.d.Id;
            saveMetadataToList(savedItemId, itemProperties, fileUploadControlID);
        });
        getItem.fail(onError);
    }).catch(addFileToFolderError => {
        console.log(addFileToFolderError);
    });
}

function getListItem(fileListItemUri) {
    // Send the request and return the response.
    return jQuery.ajax({
        url: fileListItemUri,
        type: "GET",
        headers: { "accept": "application/json;odata=verbose" }
    });
}
function onError(error) {
    alert(error.responseText);
}

var FileUploadService = function () {
    this.siteUrl = _spPageContextInfo.webAbsoluteUrl;
    this.siteRelativeUrl = _spPageContextInfo.webServerRelativeUrl != "/" ? _spPageContextInfo.webServerRelativeUrl : "";
    this.fileUpload = function (file, documentLibrary, fileName) {
        return new Promise((resolve, reject) => {
            this.createDummyFile(fileName, documentLibrary).then(result => {
                let fr = new FileReader();
                let offset = 0;
                // the total file size in bytes...  
                let total = file.size;
                // 1MB Chunks as represented in bytes (if the file is less than a MB, seperate it into two chunks of 80% and 20% the size)...  
                let length = parseInt(1000000) > total ? Math.round(total * 0.8) : parseInt(1000000);
                let chunks = [];
                //reads in the file using the fileReader HTML5 API (as an ArrayBuffer) - readAsBinaryString is not available in IE!  
                fr.readAsArrayBuffer(file);
                fr.onload = (evt) => {
                    while (offset < total) {
                        //if we are dealing with the final chunk, we need to know...  
                        if (offset + length > total) {
                            length = total - offset;
                        }
                        //work out the chunks that need to be processed and the associated REST method (start, continue or finish)  
                        chunks.push({
                            offset,
                            length,
                            method: this.getUploadMethod(offset, length, total)
                        });
                        offset += length;
                    }
                    for (var i = 0; i < chunks.length; i++) {
                        console.log(chunks[i]);
                    }
                    //each chunk is worth a percentage of the total size of the file...  
                    const chunkPercentage = (total / chunks.length) / total * 100;
                    if (chunks.length > 0) {
                        //the unique guid identifier to be used throughout the upload session  
                        const id = this.guid();
                        //Start the upload - send the data to S  
                        this.uploadFile(evt.target.result, id, documentLibrary, fileName, chunks, 0, 0, chunkPercentage, resolve, reject);
                    }
                };
            })
        });
    }
    this.createDummyFile = function (fileName, masterLibraryName) {
        return new Promise((resolve, reject) => {
            var serverRelativeUrlToFolder;
            var productCode = existingProductType === "Select" || existingProductType === undefined ? $("#ddlProductType option:selected").text() : existingProductType;
            if (productCode === "Select") {
                serverRelativeUrlToFolder = "decodedurl='" + this.siteRelativeUrl + "/" + masterLibraryName + "'";
            } else {
                serverRelativeUrlToFolder = "decodedurl='" + this.siteRelativeUrl + "/" + masterLibraryName + "/" + productCode + "'";
            }
            var endpoint = this.siteUrl + "/_api/Web/GetFolderByServerRelativePath(" + serverRelativeUrlToFolder + ")/files" + "/add(overwrite=true,url='" + fileName + "')"
            const headers = {
                "accept": "application/json;odata=verbose"
            };
            this.executeAsync(endpoint, this.convertDataBinaryString(2), headers).then(file => resolve(true)).catch(err => reject(err));
        });
    }

    // Base64 - this method converts the blob arrayBuffer into a binary string to send in the REST request  
    this.convertDataBinaryString = function (data) {
        let fileData = '';
        let byteArray = new Uint8Array(data);
        for (var i = 0; i < byteArray.byteLength; i++) {
            fileData += String.fromCharCode(byteArray[i]);
        }
        return fileData;
    }

    this.executeAsync = function (endPointUrl, data, requestHeaders) {
        return new Promise((resolve, reject) => {
            // using a utils function we would get the APP WEB url value and pass it into the constructor...  
            let executor = new SP.RequestExecutor(this.siteUrl);
            // Send the request.  
            executor.executeAsync({
                url: endPointUrl,
                method: "POST",
                body: data,
                binaryStringRequestBody: true,
                headers: requestHeaders,
                success: offset => resolve(offset),
                error: err => reject(err.responseText)
            });
        });
    }

    //this method sets up the REST request and then sends the chunk of file along with the unique indentifier (uploadId)  
    this.uploadFileChunk = function (id, libraryPath, fileName, chunk, data, byteOffset) {
        return new Promise((resolve, reject) => {
            let offset = chunk.offset === 0 ? '' : ',fileOffset=' + chunk.offset;
            //parameterising the components of this endpoint avoids the max url length problem in SP (Querystring parameters are not included in this length)  
            var productCode = existingProductType === "Select" || existingProductType === undefined ? $("#ddlProductType option:selected").text() : existingProductType;
            if (productCode === "Select") {
                productCode = "";
            }
            let endpoint = this.siteUrl + "/_api/web/getfilebyserverrelativeurl('" + this.siteRelativeUrl + "/" + libraryPath + "/" + productCode + "/" + fileName + "')/" + chunk.method + "(uploadId=guid'" + id + "'" + offset + ")";
            const headers = {
                "Accept": "application/json; odata=verbose",
                "Content-Type": "application/octet-stream"
            };
            this.executeAsync(endpoint, data, headers).then(offset => resolve(offset)).catch(err => reject(err));
        });
    }
    //the primary method that resursively calls to get the chunks and upload them to the library (to make the complete file)  
    this.uploadFile = function (result, id, libraryPath, fileName, chunks, index, byteOffset, chunkPercentage, resolve, reject) {
        //we slice the file blob into the chunk we need to send in this request (byteOffset tells us the start position)  
        const data = this.convertFileToBlobChunks(result, byteOffset, chunks[index]);
        //upload the chunk to the server using REST, using the unique upload guid as the identifier  
        this.uploadFileChunk(id, libraryPath, fileName, chunks[index], data, byteOffset).then(value => {
            const isFinished = index === chunks.length - 1;
            index += 1;
            //More chunks to process before the file is finished, continue  
            if (index < chunks.length) {
                this.uploadFile(result, id, libraryPath, fileName, chunks, index, byteOffset, chunkPercentage, resolve, reject);
            } else {
                resolve(value);
            }
        }).catch(err => {
            console.log('Error in uploadFileChunk! ');
            popUpRequestEnded();
            alert('Something went wrong, please upload the file again.');
            reject(err);
        });
    }
    //Helper method - depending on what chunk of data we are dealing with, we need to use the correct REST method...  
    this.getUploadMethod = function (offset, length, total) {
        if (offset + length + 1 > total) {
            return 'finishupload';
        } else if (offset === 0) {
            return 'startupload';
        } else if (offset < total) {
            return 'continueupload';
        }
        return null;
    }
    //this method slices the blob array buffer to the appropriate chunk and then calls off to get the BinaryString of that chunk  
    this.convertFileToBlobChunks = function (result, byteOffset, chunkInfo) {
        let arrayBuffer = result.slice(chunkInfo.offset, chunkInfo.offset + chunkInfo.length);
        return this.convertDataBinaryString(arrayBuffer);
    }
    this.guid = function () {
        function s4() {
            return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1);
        }
        return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
    }
}

