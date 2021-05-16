var libraryName = "eSource Content Library";
var adminGroupName = "RFG eSource Content Management Owners";
var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.webServerRelativeUrl;
google.charts.load('current', {
	'packages': ['corechart']
});

checkCurentUserInSPGroup(adminGroupName).then(function (data) {
    if (data.d.results.length > 0) {
        google.charts.setOnLoadCallback(drawChart);
    } else {
        alert("You are not authorised to view the dashboard as you are not in the Admin group of the site.")
    }
});

function getAllItemsFromLib(apiPath, results, deferred) {
	var results = results || [];
	results.data = results.data || [];
	var deferred = deferred || $.Deferred();
	
	$.ajax({
		url: apiPath,
		headers: {
			Accept: "application/json;odata=verbose"
		},
		async: false,
		success: function (data) {
			if (results.data.length == 0) {
				results.data = data.d.results;
			}else {
				results.data = results.data.concat(data.d.results);
			}
			if(data.d.__next) {
				apiPath = data.d.__next;
				getAllItemsFromLib(apiPath, results, deferred);
			}else {
				deferred.resolve(results);
				console.log(results);
			}
		},
		error: function (data) {
			console.log("An error occurred. Please try again.");
		}
	});
	return deferred.promise();
}

function drawChart() {
	var apiPath = siteURL + "/_api/web/lists/getByTitle('" + libraryName + "')/Items?$select=ContentType/Id,ContentType/Name,DocumentStatus&$expand=ContentType&$top=5000";
	getAllItemsFromLib(apiPath).then(function(results) {
		prepareDataForContentType(results.data);
		prepareDataForDocumentStatus(results.data);
	});
}

function prepareDataForContentType(resultsColl) {
	var statusDataAfterFilter = [], itemArray = [], listData;
	listData = new google.visualization.DataTable();
	listData.addColumn('string', 'Content Types');
	listData.addColumn('number', 'Count');
		statusCountData = _.countBy(resultsColl, function (value) {
		return value["ContentType"].Name;
	});
	var totalValue = 0;
	$.each(statusCountData, function (key, value) {
		totalValue = totalValue + value;
	});
	$.each(statusCountData, function (key, value) {
		listData.addRow([key, parseInt(value)]);
		statusDataAfterFilter.push(key);
		itemArray.push(new StatusObject(key, value, totalValue))
	});
	itemArray.sort(function (a, b) {
		if (a.ContentType.Name < b.ContentType.Name) return -1;
		if (a.ContentType.Name > b.ContentType.Name) return 1;
		return 0;
	});
	drawChartForContentType(listData);	
}

function prepareDataForDocumentStatus(resultsColl) {
	var statusDataAfterFilter = [], itemArray = [], listData_DocStatus;
	listData_DocStatus = new google.visualization.DataTable();
	listData_DocStatus.addColumn('string', 'Document Status');
	listData_DocStatus.addColumn('number', 'Count');	
	statusCountData = _.countBy(resultsColl, function (value) {
		if (value.DocumentStatus !== null && value.DocumentStatus !== "undefined") {
			return value["DocumentStatus"];
		}		
	});
	var totalValue = 0;
	$.each(statusCountData, function (key, value) {
		if (key !== "undefined") {
			totalValue = totalValue + value;
		}
		
	});
	$.each(statusCountData, function (key, value) {
		if (key !== "undefined") {
			listData_DocStatus.addRow([key, parseInt(value)]);
			statusDataAfterFilter.push(key);
			itemArray.push(new StatusObject(key, value, totalValue))
		}
	});
	itemArray.sort(function (a, b) {
		if (a.DocumentStatus < b.DocumentStatus) return -1;
		if (a.DocumentStatus > b.DocumentStatus) return 1;
		return 0;
	});
	drawChartForDocumentStatus(listData_DocStatus);	
}

function drawChartForContentType(listData) {
	var siteStatusCountWise,
		options = {
			title: "Content Type",
			colors: ['#5291DD', '#F58144', '#FCB13B', '#EC9A1E', '#E74A21'],
			fontSize: 15,
			legendFontSize: 15,
			titleFontSize: 16,
			tooltipFontSize: 15
		};
	siteStatusCountWise = new google.visualization.PieChart(document.getElementById('piechart'));
	siteStatusCountWise.draw(listData, options);
}

function drawChartForDocumentStatus(listData) {
	var siteStatusCountWise,
		options = {
			title: "Document Status",
			colors: ['#5291DD', '#F58144', '#FCB13B', '#EC9A1E', '#E74A21'],
			fontSize: 15,
			legendFontSize: 15,
			titleFontSize: 16,
			tooltipFontSize: 15
		};
	siteStatusCountWise = new google.visualization.PieChart(document.getElementById('piechart_DocStatus'));
	siteStatusCountWise.draw(listData, options);
}

function StatusObject(contentType, count, totalValue) {
	if (contentType) {
		this.ContentType = contentType;
	}
	if (count) {
		this.Count = count;
	}
	if (count == '0') {
		this.SiteCountInPer = '0%';
	}
}

function checkCurentUserInSPGroup(groupName) {
    var apiPath = siteURL + "/_api/web/sitegroups/getByName('" + groupName + "')/Users?$filter=Id eq " + _spPageContextInfo.userId;
    return $.ajax({
		url: apiPath,
        headers: {
            Accept: "application/json;odata=verbose"
        },
        async: false
    });										  
}

