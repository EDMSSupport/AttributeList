/*
* Created by: Vasanth
* Created Date: 06/12/2017
* Description: script for CSV File Generation
* Functional Impact: Download CSV Button click , Initial Function Load , Search Result Generation
*/
(function () {

    var EDMS = window.EDMS || {};

    /* properties */
    var CreatingList = function () {
        this.headerRow = [];
        this.managedProperty = [];
        this.csvData = [];
        this.queryText = "";
        this.refiner = [];
        this.sortName = "";
        this.sortOrder = null;
        this.DatesFields = ["ResearchStartDate", "ResearchEndDate", "ApprovalDate", "DateEntered", "ContactDate", "MeetingDate", "RetentionDate", "ApprovedDate"];
        this.ENGLISHSEARCH = ["TitleEnglish", "OverviewEnglish", "ReportTypeOtherEnglish", "AuthorNameEnglish", "CommentEnglish",
							"AgreementMatterEnglish", "ObjectiveEnglish", "AgendaEnglish"];
        this.JAPANESESEARCH = ["TitleJapanese", "OverviewJapanese", "ReportTypeOtherJapanese", "AuthorNameJapanese", "CommentJapanese",
							"AgreementMatterJapanese", "ObjectiveJapanese", "AgendaJapanese"];
		this.IGNORESEARCH = ["Function1stDigit","Part1stDigit","ENGFamily","MeetingType","ObjectiveAgreement","ObjectiveAgreementMatter","Comment","CadicsCategoryName","FileType"];
		this.MODIFYQUERYPARAM = ["AgreementMatterEnglish1","AgreementMatterJapanese1","ObjectiveJapanese1","ObjectiveEnglish1"]
		//this.MODIFYQUERYCOMMENTPARAM = ["NissanSubmittedDocument1", "ContactPartnerSubmittedDocument1", "MeetingDetail1"]
		this.PEOPLEPICKER = ["NMLPersonName","AuthorNameID"]
		this.MODIFYTITLEQUERYPARAM = ["TitleEnglish1","TitleJapanese1"]
		this.COMMASEPVALUES1 = ["ContactSecID","ContactName","ContactPosition","Drawperson"];
		this.COMMASEPVALUES2 = ["ContactUserID","Contacts"];
		this.count = {};
    };

    /* method */
    CreatingList.prototype = {

        /*
        * Created by: Vasanth
        * Created Date: 06/12/2017
        * Description: Function used to create CSV File Object
        * Functional Impact: CSV File Generation
        */

        eventDownload: function (csvArray, fileName) {

            var csv = csvArray.join('\n');

            var mimeType = 'text/plain';
            fileName = fileName + '.csv';

            var bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
            var blob = new Blob([bom, csv], { type: mimeType });

            var downloadAnchor = document.createElement('a');
            downloadAnchor.download = fileName;
            downloadAnchor.target = '_blank';

            if (window.navigator.msSaveBlob) {
                window.navigator.msSaveBlob(blob, fileName)
            }
            else if (window.URL && window.URL.createObjectURL) {
                downloadAnchor.href = window.URL.createObjectURL(blob);
                document.body.appendChild(downloadAnchor);
                downloadAnchor.click();
                document.body.removeChild(downloadAnchor);
            }
            else if (window.webkitURL && window.webkitURL.createObject) {
                downloadAnchor.href = window.webkitURL.createObjectURL(blob);
                downloadAnchor.click();
            }
            else {
                window.open('data:' + mimeType + ';base64,' + window.Base64.encode(csv), '_blank');
            }
            /* loading screen is removed */
            $('.ms-dlgOverlay').remove();
        },
        /*
        * Created by: Vasanth
        * Created Date: 13/07/2018
        * Description: Function used to handle comma, newline etc., characters in the string
        * Functional Impact: CSV File Generation
        */
		stringToCSVCell : function(str)
        {
            var mustQuote = (str.contains(",") || str.contains("\"") || str.contains("\r") || str.contains("\n"));
            if (mustQuote)
            {
                var sb = "\"";
				for (var charIndex = 0; charIndex < str.length; charIndex++)
                {
                    sb+=str.charAt(charIndex);
                    if (str.charAt(charIndex) == '"')
                        sb+="\"";
                }
                sb+="\"";
                return sb.toString();
            }

            return str;
        },

        /*
        * Created by:Vasanth
        * Created Date:12/14/2017
        * Description:function used to create search result query
        * Functional Impact: Search Result Generation
        * Modified by:<Author>
        * Modified Date:<Date>
        * Modified Reason:<What for?>
        */
		
        customSearch: function (startRowIndex) {

            var deferredObject = $.Deferred();
            /* Waits the function to be executed until SP.js , SP.ClientContext are loaded */
            SP.SOD.executeFunc("SP.js", "SP.ClientContext", function () {
                /* Waits the function to be executed until SP.Search.js , Microsoft.SharePoint.Client.Search.Query.KeywordQuery are loaded */
                SP.SOD.executeFunc("SP.Search.js", "Microsoft.SharePoint.Client.Search.Query.KeywordQuery", function () {

                    /* Create the Context */
                    var context = new SP.ClientContext(_spPageContextInfo.siteServerRelativeUrl);

                    /* Creating the KeywordQuery Object */
                    var keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(context);

                    /* Setting the query text to be fetched by the search query */
                    keywordQuery.set_queryText(EDMS.creatingList.queryText);
                    keywordQuery.set_ignoreSafeQueryPropertiesTemplateUrl(false);

                    /* Setting the managed property to be fetched by the search query */
                    var properties = keywordQuery.get_selectProperties();
                    for (var count = 0; count < EDMS.creatingList.managedProperty.length; count++) {
                        if (EDMS.creatingList.managedProperty[count] !== "") {
                            properties.add(EDMS.creatingList.managedProperty[count]);
                        }
                    }

                    /* setting the refiner filter to be fetched by the search query */
                    if (EDMS.creatingList.refiner.length > 0) {
                        var filterCollection = keywordQuery.get_refinementFilters();
                        for (var refiners = 0; refiners < EDMS.creatingList.refiner.length; refiners++) {
                            filterCollection.add(EDMS.creatingList.refiner[refiners]);
                        }
                    }

                    /* setting the sorting filter to be fetched by the search query */
                    if (EDMS.creatingList.sortName != "" && EDMS.creatingList.sortOrder != null) {
                        keywordQuery.set_enableSorting(true);
                        var sortproperties = keywordQuery.get_sortList();
                        sortproperties.add(EDMS.creatingList.sortName, EDMS.creatingList.sortOrder);
                    }

                    /* Setting the row limit and the start row index to be fetched by the search query */
                    keywordQuery.set_rowLimit(500);
                    keywordQuery.set_startRow(startRowIndex);

                    /* Set Trim Duplicates to False */
                    keywordQuery.set_trimDuplicates(false);
 					/* Set Culture to 17 */
                    keywordQuery.set_culture(17);

                    /* Set the query time out 0-60000 */
                    keywordQuery.set_timeout(60000);

                    /* Creating Search Executor Object */
                    var searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(context);

                    /* Executing the Search Query */
                    var results = searchExecutor.executeQuery(keywordQuery);

                    /* Executing the Context */
                    context.executeQueryAsync(onQueryListSuccess, onQueryFail)

                    /* function used when the search query execution is successful */
                    function onQueryListSuccess() {
                        /*traverses through the search result row by row */
                        if (results.m_value.ResultTables.length > 0) {
                            for (var rowCount = 0; rowCount < results.m_value.ResultTables[0].RowCount; rowCount++) {
                                var row = [];
                                var searchRow = results.m_value.ResultTables[0].ResultRows[rowCount];
                                /* traverses through each managed property */
                                for (var propertyCount = 0; propertyCount < EDMS.creatingList.managedProperty.length; propertyCount++) {
                                    /* checks for the managed property and push the managed property value to the temporary row */
                                    var managePropertyValue = String(searchRow[EDMS.creatingList.managedProperty[propertyCount]] ? searchRow[EDMS.creatingList.managedProperty[propertyCount]] : "");
                                    
                                  	 if(EDMS.creatingList.managedProperty[propertyCount]=="Author")
                                  	 {
                                  	 managePropertyValue = managePropertyValue.split(';')[0];
                                  	 }
								    if(managePropertyValue =='・' || managePropertyValue.toLowerCase() =='<blank>')
									{
									managePropertyValue = managePropertyValue.replace(managePropertyValue,'');
									}
                                    if ((EDMS.creatingList.PEOPLEPICKER.indexOf(EDMS.creatingList.managedProperty[propertyCount]) !== -1) && managePropertyValue.indexOf('|') !== -1) {
                                        managePropertyValue = managePropertyValue.split('|')[0].replace(/\ /g,"") == "" ? managePropertyValue.split('|')[1] : managePropertyValue.split('|')[0];										
                                    }
									if (EDMS.creatingList.COMMASEPVALUES1.indexOf(EDMS.creatingList.managedProperty[propertyCount]) !== -1) {
                                       var managePropertyValue1 = managePropertyValue.split('$#$');
									   var managePropertyValue2 = "";
									   $.each(managePropertyValue1,function(index){
									   if(managePropertyValue1[index]!="")
									   {
									   if(managePropertyValue2 == "")
									   managePropertyValue2=managePropertyValue1[index];
									   else
									   {
									   managePropertyValue2 = managePropertyValue2 + ';' +  managePropertyValue1[index];
									   }
									   }
									   });
									   managePropertyValue = managePropertyValue2;
									
                                    }
									if(EDMS.creatingList.managedProperty[propertyCount]=="InternalTelephoneNumber")
									{
									if(managePropertyValue!='・' || managePropertyValue.replace(/\ /g, "")!='')
									managePropertyValue= "'" + managePropertyValue
									}
									if (EDMS.creatingList.COMMASEPVALUES2.indexOf(EDMS.creatingList.managedProperty[propertyCount]) !== -1) {
                                       var managePropertyValue3 = managePropertyValue.split('$$#$$');
									   var managePropertyValue4 = "";
									   $.each(managePropertyValue3,function(index){
									   if(managePropertyValue3[index]!="-" && managePropertyValue3[index]!="")
									   {
									   if(managePropertyValue4 == "")
								       managePropertyValue4 = managePropertyValue3[index];
									   else
									   managePropertyValue4 = managePropertyValue4 + ';' + managePropertyValue3[index];
									   }
									   }); 
									   managePropertyValue = managePropertyValue4.split('$#$').join(';');
                                    }
                                    if (EDMS.creatingList.DatesFields.indexOf(EDMS.creatingList.managedProperty[propertyCount]) != -1) {
                                        var newDateFormat = "";
                                        if (managePropertyValue != "" && managePropertyValue != undefined && managePropertyValue != null) {
                                            var dtValue = new Date(managePropertyValue);
                                            var day = dtValue.getDate();
                                            day = day.toString().length > 1 ? day : '0' + day;
                                            var month = (dtValue.getMonth() + 1);
                                            month = month.toString().length > 1 ? month : '0' + month;
                                            newDateFormat = dtValue.getFullYear() + '/' + month + '/' + day;
                                        }
                                        managePropertyValue = newDateFormat;
                                    }
									managePropertyValue = EDMS.creatingList.stringToCSVCell(managePropertyValue);
                                    if (EDMS.creatingList.managedProperty[propertyCount] == 'Path') {
									var appName= managePropertyValue.split('/');
									var Title = managePropertyValue.split('/')[appName.length-1];
									var linkSite = managePropertyValue.split('/'+Title)[0];
									linkSite = linkSite.substring(0,linkSite.lastIndexOf("/"));
									linkSite = linkSite + '/_layouts/15/DocSetHome.aspx?id=' + managePropertyValue.split(location.host)[1];
									managePropertyValue = linkSite;
                                    }
                                    /* managed property value are inserted to a temporary array */
                                    row.push(managePropertyValue);
                                }
                                /* the managed property values for each search result row is pushed to the array for CSV */
                                EDMS.creatingList.csvData.push(row);

                            }
                            /* checks the TotalRows count for the search query result with the RowCount count retrieved by the search query result */
                            if (results.m_value.ResultTables[0].TotalRows <= startRowIndex + results.m_value.ResultTables[0].RowCount) {
                                /* calls function to create CSV File for the search result */
                                EDMS.creatingList.eventDownload(EDMS.creatingList.csvData, 'SearchResult');
                            }
                            else {
                                /* loops again until all results are fetched */
                                EDMS.creatingList.customSearch(startRowIndex + results.m_value.ResultTables[0].RowCount);
                            }
                        }
                        else {
                            $('.ms-dlgOverlay').remove();

                        }

                    }
                    /* function used when the search query is failed */
                    function onQueryFail(sender, args) {
                        alert('Query failed. Error:' + args.get_message());
                        /* loading screen is removed */
                        $('.ms-dlgOverlay').remove();
                    }
                });
            });

            return deferredObject.promise();
        },

        /*
        * Created by:Vasanth
        * Created Date:06/12/2017
        * Description:initiate function is called after the script is loaded
        * Functional Impact: Initial Function Call
        */

        init: function () {

            var _self = this;

            /*
            * Created by:Vasanth
            * Created Date:06/12/2017
            * Description:function for Download List button
            * Functional Impact: Download list button click
            */

            $(document).on('click', '.edms-down', function () {

                var row = [];
                EDMS.creatingList.csvData = [];
                EDMS.creatingList.headerRow = [];
                EDMS.creatingList.managedProperty = [];
                EDMS.creatingList.refiner = [];
                EDMS.creatingList.queryText = "";
                EDMS.creatingList.sortName = "";
                EDMS.creatingList.sortOrder = null;

                /* loading screen is shown at the start of the function */
                if (!$('.ms-dlgOverlay').length) {
                    $('body').append('<div class="ms-dlgOverlay edms-loading" style="display:block"></div>');
                }

                /* creating query text for the search query */
                _self.createQueryText();

                /* creating refiner filter and sorting filter for the search query */
                _self.createRefinerSortFilter();                
                
                /* retrieving common attributes and their managed property for the search query */
                //_self.createAttributes(downloadCSVPrep[localStorage.getItem(EDMS.common.APPNAME_KEY)]);				
                _self.createAttributes(downloadCSVPrep[GetUrlKeyValue(EDMS.common.APPNAME).toUpperCase().replace(/\ /g, "")]);
                _self.createAttributes(downloadCSVPrep.COMMON);

                /* retrieving specific attributes and their managed property for each Type of Standard values for the search query */
                _self.createSpecificAttributes();

                /* removes duplicate values on headerRow array */
                EDMS.creatingList.headerRow = EDMS.creatingList.headerRow.removeDuplicates();

                /* removes duplicate values on managedProperty array */
                EDMS.creatingList.managedProperty = EDMS.creatingList.managedProperty.removeDuplicates();

                /* setting the header row for the CSV */
                for (var count = 0; count < EDMS.creatingList.headerRow.length; count++) {
                    if (EDMS.creatingList.headerRow[count] !== "" && EDMS.creatingList.managedProperty[count] !== "") {
                        row.push(EDMS.creatingList.headerRow[count]);
                    }
                }
                EDMS.creatingList.csvData.push(row);

                /* calls client search query function */
                _self.customSearch(0);

            });
/* sorting functions are called for the table header row of the search result */
$(document).on('click', 'th', function (e) {
e.stopPropagation();
//e.preventDefault();
var sort="";
var lang = localStorage.getItem(EDMS.common.LANGUAGE_KEY) || 0;
var currentSort=$(this).attr('data-wordid');
if(currentSort!='Common_Document')
sort = sortProperties[currentSort][lang];
if(EDMS.creatingList.count[sort] == 1)
{
EDMS.creatingList.count[sort]=0;
$getClientControl(this).sortOrRank(sort+'_DES');
}
else
{
EDMS.creatingList.count[sort]=1;
$getClientControl(this).sortOrRank(sort+'_ASC');
}

return false;
});

        },
        /*
        * Created by:Vasanth
        * Created Date:06/12/2017
        * Description:function to set search query text
        * Functional Impact:search query text setting for search query
        */
        createQueryText: function () {
			EDMS.creatingList.queryText = "";
            var queryDateApplication = [];
            var queryApprovalDate = [];
			var queryApprovedDate = [];
            var queryRetentionDate = [];
            var queryMeetingDate = [];
            var queryContactDate = [];
            var queryResearchStartDate = [];
            var queryResearchEndDate = [];
			var queryCreated = false;
			var queryCreated2 = false;
			var tarDoc = GetUrlKeyValue(EDMS.common.APPNAME).toUpperCase().replace(/\ /g,"");
			
			//Default sorting if refinary not used
			// Sorting (Ascending = 0, Descending = 1)
			if(tarDoc == "MM")
			{				
                EDMS.creatingList.sortOrder = 1;
                EDMS.creatingList.sortName = "LastModifiedTime";
			}
			
			if(tarDoc == "ED")
			{
                EDMS.creatingList.sortOrder = 1;
                EDMS.creatingList.sortName = "ApprovedDate";
			}
			if(tarDoc == "RCR")
			{
                EDMS.creatingList.sortOrder = 1;
                EDMS.creatingList.sortName = "ApprovalDate";
			}
			
            /* checks the URL to get query string only */
            var queryString = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
            for (var queryCount = 0; queryCount < queryString.length; queryCount++) {
                var queryParameterValueSplit = queryString[queryCount].split('=');
                var queryParameter = decodeURIComponent(queryParameterValueSplit[0]);
                var queryValue = decodeURIComponent(queryParameterValueSplit[1]);

                queryValue = queryString[queryCount].contains("#k=#s=") ? "" : queryValue;
                
                if(EDMS.creatingList.MODIFYTITLEQUERYPARAM.indexOf(queryParameter) != -1)
                {
                	queryParameter = queryParameter.slice(0,-1)
                }
                
                if(queryValue.match(/\s/g))
                {
                queryValue = '"'+queryValue+'"';
                }
                
                /* appends to the search query text with managed property and query value */ 
                               
                if (queryValue != "" && queryParameter != "" && queryParameter.match(/\d+/g) == null && queryParameter.indexOf('#Default') === -1
                    && queryParameter != "ResearchStartDate" && queryParameter != "ResearchEndDate" 
                    && EDMS.creatingList.JAPANESESEARCH.indexOf(queryParameter) == -1 && EDMS.creatingList.IGNORESEARCH.indexOf(queryParameter) == -1
                    && EDMS.creatingList.MODIFYQUERYPARAM.indexOf(queryParameter) == -1) {
                    if (EDMS.creatingList.ENGLISHSEARCH.indexOf(queryParameter) != -1) {
                        var jpQuery = EDMS.creatingList.JAPANESESEARCH[EDMS.creatingList.ENGLISHSEARCH.indexOf(queryParameter)];
                        EDMS.creatingList.queryText += queryParameter + ":" + queryValue + " OR " + jpQuery + ":" + queryValue + " ";
                    }
                    else {
                       // (/\s/g.test(queryValue) && queryParameter !== "EDMSTypeofstandard") ? EDMS.creatingList.queryText += queryParameter + ":\"" + queryValue + "\" " :
                                                    if(queryParameter=="Meetingname")
													EDMS.creatingList.queryText += queryParameter + "=" + queryValue + " ";
													else
													EDMS.creatingList.queryText += queryParameter + ":" + queryValue + " ";
                    }
                    
                }
				
				if(EDMS.creatingList.MODIFYQUERYPARAM.indexOf(queryParameter) != -1 && !queryCreated)
				{
					$.each(EDMS.creatingList.MODIFYQUERYPARAM,function(index){
						var currentItem = this.slice(0,-1)
						if(index < EDMS.creatingList.MODIFYQUERYPARAM.length-1)
						  EDMS.creatingList.queryText += currentItem + ":" + queryValue + " OR "
						else
						  EDMS.creatingList.queryText += currentItem + ":" + queryValue + " "
					});
					queryCreated = true;
				}
                /* checks the query parameter for any numerical digits */
                if (queryParameter.match(/\d+/g) != null || queryParameter == "ResearchStartDate" || queryParameter == "ResearchEndDate") {
                    /* appends to the temporary array to get both From Date and To Date for each type of Date respectively */
                    if (queryParameter.indexOf('ApprovalDate1') !== -1 || queryParameter.indexOf('ApprovalDate2') !== -1) {
                        queryApprovalDate.push(queryValue);
                    }
					if (queryParameter.indexOf('ApprovedDate1') !== -1 || queryParameter.indexOf('ApprovedDate2') !== -1) {
                        queryApprovedDate.push(queryValue);
                    }
                    if (queryParameter.indexOf('RetentionDate1') !== -1 || queryParameter.indexOf('RetentionDate2') !== -1) {
                        queryRetentionDate.push(queryValue);
                    }
                    if (queryParameter.indexOf('MeetingDate1') !== -1 || queryParameter.indexOf('MeetingDate2') !== -1) {
                        queryMeetingDate.push(queryValue);
                    }
                    if (queryParameter.indexOf('ContactDate1') !== -1 || queryParameter.indexOf('ContactDate2') !== -1) {
                        queryContactDate.push(queryValue);
                    }
                    if (queryParameter.indexOf('ResearchStartDate') !== -1 || queryParameter.indexOf('ResearchEndDate') !== -1) {
                        queryResearchStartDate.push(queryValue);
                        queryResearchEndDate.push(queryValue);
                    }
                    /*if (queryParameter === 'ResearchStartDate') {
                        EDMS.creatingList.queryText += queryParameter + ">" + queryValue + " ";
                    }
                    else if (queryParameter === 'ResearchEndDate') {
                        EDMS.creatingList.queryText += queryParameter + "<" + queryValue + " ";
                    }*/
                    /*if (queryParameter.indexOf('ResearchStartDate') !== -1 || queryParameter.indexOf('ResearchEndDate') !== -1) {
                        queryDateApplication.push(queryValue);
                    }*/
                }
            }
            /* appends to the search query text when each type of Date has both From Date and To Date respectively */
            if (queryApprovalDate.length == 2)
                EDMS.creatingList.queryText += "ApprovalDate:" + queryApprovalDate.join("..") + " ";
			if (queryApprovedDate.length == 2)
                EDMS.creatingList.queryText += "ApprovedDate:" + queryApprovedDate.join("..") + " ";
            if (queryDateApplication.length == 2)
                EDMS.creatingList.queryText += "ResearchStartDate:" + queryDateApplication.join("..") + " ";
            if (queryRetentionDate.length == 2)
                EDMS.creatingList.queryText += "RetentionDate:" + queryRetentionDate.join("..") + " ";
            if (queryMeetingDate.length == 2)
                EDMS.creatingList.queryText += "MeetingDate:" + queryMeetingDate.join("..") + " ";
            if (queryContactDate.length == 2)
                EDMS.creatingList.queryText += "ContactDate:" + queryContactDate.join("..") + " ";
            if (queryResearchStartDate.length == 2 && queryResearchEndDate.length == 2)
            {
                EDMS.creatingList.queryText += "ResearchStartDate:" + queryResearchStartDate.join("..") + " OR ";
                EDMS.creatingList.queryText += "ResearchEndDate:" + queryResearchEndDate.join("..") + " ";
            }

            /* appends to the search query text with static managed property condition */
            //EDMS.creatingList.queryText += 'ContentType:EDMS_DocumentSet EDMSStatus:"Issued" OR EDMSStatus:"Abolished"';
            tarDoc = tarDoc == "MMM" ? "MM" : tarDoc;
			EDMS.creatingList.queryText += 'ContentType:EDMS_'+tarDoc+'_DocumentSet EDMSStatus:"Issued"';
        },
        /*
        * Created by:Vasanth
        * Created Date:06/12/2017
        * Description:function to set refiner and sort in the search result query
        * Functional Impact:setting for refiners and sort on the search query result if refiners or sorting is present
        */

        createRefinerSortFilter: function () {
            /* checking for any refiners and sorting setting on the search result */
            var refinerCheck = null;
            var tempClientControlID = localStorage.getItem("Client_ControlID");
            var clientControlIDURL = "#" + tempClientControlID + "=";
            /* normal query string parameter when sorting or refiner is used - Default */
            if (window.location.href.indexOf("#Default=") !== -1) {
                refinerCheck = window.location.href.split("#Default=");
            }
            /* query string parameter when sorting or refiner is used with clientControlID*/
            if (window.location.href.indexOf(clientControlIDURL) !== -1) {
                refinerCheck = window.location.href.split(clientControlIDURL);
            }
            /* retrieving particular query string value */
            if (refinerCheck != null) {
                var refinerString = decodeURIComponent(refinerCheck[1]);
                var hasParseError = false;
                var refinerJSON;
                try {
                    refinerJSON = JSON.parse(refinerString.replace(/\\/g, "").replace(/""/g, "\"").replace(/"k":"/g, '"k":""'))
                }
                catch (err) {
                    refinerJSON = JSON.parse(refinerString);
                    hasParseError = true;
                }
                /* retrieves the refiner property used on the search result */
                if (refinerJSON.r != null) {
                    for (var refinerCount = 0; refinerCount < refinerJSON.r.length; refinerCount++) {
                        var refinerManagedPropertyName = refinerJSON.r[refinerCount].n;
                        var refinerToken = refinerJSON.r[refinerCount].t[0];
                        if (EDMS.creatingList.DatesFields.indexOf(refinerManagedPropertyName) != -1) {
                            EDMS.creatingList.refiner.push(refinerManagedPropertyName + ":" + refinerToken);
                        }
                        else {
                            refinerToken = hasParseError ? refinerToken.replace(/"/g, '') : refinerToken;
                            EDMS.creatingList.refiner.push(refinerManagedPropertyName + ":\"" + refinerToken + "\"");
                        }
                    }
                }
                /* retrieves the sort property used on the search result */
                if (refinerJSON.o != null) {
                    for (var sortCount = 0; sortCount < refinerJSON.o.length; sortCount++) {
                        EDMS.creatingList.sortOrder = refinerJSON.o[sortCount].d;
                        EDMS.creatingList.sortName = refinerJSON.o[sortCount].p;
                    }
                }
            }
        },

        /*
        * Created by:Vasanth
        * Created Date:06/12/2017
        * Description:function to set header row of CSV and managed properties of their attributes in the search query
        * Functional Impact:setting for header row and managed properties for search query
        */

        createAttributes: function (downloadAttribute) {
            /* check for current lanaguage */
            var stLang = Number(JSON.parse(localStorage.getItem(EDMS.common.LANGUAGE_KEY)));
            var lang = isFinite(stLang) ? stLang : 0;

            /* retrieving the attribute names and their managed property */
            for (var count = 0; count < downloadAttribute.length; count++) {
                if (downloadAttribute[count].head !== "" && downloadAttribute[count].prop !== "") {
                    /* header row text are checked with their current language from the downloadCSV JSON and pushed to the headerRow array */
                    var headerText = ""
					var managedPropertyText = ""
					var currentDataObj = downloadAttribute[count];
					if(lang == 1)
					{
						headerText = currentDataObj.headJP || currentDataObj.head;
						managedPropertyText = currentDataObj.propJP || currentDataObj.prop;
					}
					else
					{
						headerText = currentDataObj.head;
						managedPropertyText = currentDataObj.prop;
					} 
                    /* headerText is inserted to headerRow array */
                    EDMS.creatingList.headerRow.push(headerText);
                    /* managed property from the downloadCSV JSON is inserted to managed property array */
                    EDMS.creatingList.managedProperty.push(managedPropertyText);
                }
            }
        },

        getFilesinDocumentsSet: function (documentSetUrl) {
            var fileURLS = ""
            var relativeUrl = documentSetUrl.split(location.host)[1];
            var splittedStr = relativeUrl.split('/');
			var actSplit = splittedStr[splittedStr.length-2]
            var targetWeb = relativeUrl.split('/' + actSplit)[0];

            //var targetWeb = relativeUrl.split('/TestDocument')[0];
            var url = targetWeb + '/_api/web/getfolderbyserverrelativeurl(\'' + relativeUrl + '\')/Files';
            $.ajax({
                url: url,
                method: 'GET',
                async: false,
                headers: {
                    "Accept": "application/json; odata=verbose",
                    "content-type": "application/json; odata=verbose"
                    //"X-RequestDigest": document.getElementById("__REQUESTDIGEST").value
                }
            }).done(function (data) {
                var siteAbsoluteURL = _spPageContextInfo.siteAbsoluteUrl;
                $.each(data.d.results, function () {
                    fileURLS += siteAbsoluteURL + this.ServerRelativeUrl + ";";
                });
            }).error(function (e) {
                console.log('error', e);
            });
            return fileURLS;
        },

        /*
        * Created by:Vasanth
        * Created Date:06/12/2017
        * Description:function to retrieve specific attribute names and managed property for each Type of Stanadard value
        * Functional Impact:retrive the JSON object for each Type of Stanadard value consist of attribute names and managed property
        */

        createSpecificAttributes: function () {
            var typeOfStandardCollection = [];
            var typeOfStandard;

            /* getting Type of Standard values from the refiner */
            var refineType = $('.ms-ref-refiner');
            $.each(refineType, function () {
                var target = $(this);
                var refiner = target.attr('RefinerName');
                if (refiner == "EDMSTypeofstandard") {
                    /* checks for short list in Type of Standard refiner */
                    var refinerShortList = target.find('.ms-ref-unsel-shortList .ms-displayBlock');
                    if (refinerShortList.length > 0) {
                        $.each(refinerShortList, function () {
                            typeOfStandard = $(this).attr('title').split('Refine by: ')[1] || $(this).attr('title').split(': ')[1];
                            typeOfStandardCollection.push(typeOfStandard);
                        });
                    }
                    /* checks for long list in Type of Standard refiner */
                    var refinerLongList = target.find('.ms-ref-unsel-longList .ms-displayBlock');
                    if (refinerLongList.length > 0) {
                        typeOfStandardCollection = [];
                        $.each(refinerLongList, function () {
                            typeOfStandard = $(this).attr('title').split('Refine by: ')[1] || $(this).attr('title').split(': ')[1];
                            typeOfStandardCollection.push(typeOfStandard);
                        });
                    }
                    /* checks for selected list in Type of Standard refiner */
                    var refinerSelectedList = target.find('.ms-ref-selSec .ms-displayBlock');
                    if (refinerSelectedList.length > 0) {
                        typeOfStandardCollection = [];
                        $.each(refinerSelectedList, function () {
                            typeOfStandard = $(this).attr('title').split('Refine by: ')[1] || $(this).attr('title').split(': ')[1];
                            typeOfStandardCollection.push(typeOfStandard);
                        });
                    }
                }
            });


            /* traverses through the Type of Standard values in the array stored */
            for (var count = 0; count < typeOfStandardCollection.length; count++) {
                typeOfStandardCollection[count] = decodeURIComponent(typeOfStandardCollection[count]);
            }
        }
    };

    /* creates a constructor of CreatingList and calls their initial function */
    EDMS.creatingList = new CreatingList();
    EDMS.creatingList.init();
    window.EDMS = EDMS;
    /* array filter function to return array with no duplicate values */
    Array.prototype.removeDuplicates = function () {
        return this.filter(function (item, index, inputArray) {
            return inputArray.indexOf(item) == index;
        });
    };


})();

