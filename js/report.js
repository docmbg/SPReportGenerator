/* jshint -W110 */

var subsites = [],
    fileCollection = [],
    ep = new ExcelPlus(),
    sName;

//create sheet to hold the information
ep.createFile("Sheet1");

function createFile(arr, fileName) {
    //simply give the write method the 2d array as content value
    ep.write({
        "content": arr
    });
    //finally save the file
    return ep.saveAs(fileName + ".xlsx");
}

//variable to hold the static name of the Document Type column


function getCurrentSite() {
    var dfd = $.Deferred();
    dfd.resolve($().SPServices.SPGetCurrentSite());
    return dfd.promise();
}


// will use in RIST Page
function getStaticName() {
    var currentSite;

    getCurrentSite().done(function(dfdResolve) {
        currentSite = dfdResolve;
    });

    var dfd = $.Deferred();

    $SP().list("Inactive Records", currentSite).info(function(fields) {
        for (var i = 0; i < fields.length; i++) {
            if (fields[i].DisplayName === "Document Type") {

                dfd.resolve(fields[i].StaticName);

                for (var choice in fields[i].Choices) {
                    console.log(fields[i].Choices[choice]);
                    var recTypeSelect = document.getElementById("recordTypes");
                    recTypeSelect.options[recTypeSelect.options.length] =
                        new Option(fields[i].Choices[choice], fields[i].Choices[choice]);
                }
            }
        }
    });
    return dfd.promise();
}

getStaticName().done(function(a) {
    sName = a;
});

function genSitesArray() {
    console.log("Generating array of sub-sites");
    subsites = $("#subsites input:checkbox:checked").map(function() {
        return $(this).val();
    }).get();
    return $.Deferred().resolve(false);
}

$(document).ready(function() {

    $("#checkAll").change(function() {
        $("input:checkbox").prop('checked', $(this).prop("checked"));
    });

    $("#getFilesBtn").click(function(event) {
        event.preventDefault();
        genSitesArray().done(function() {
            $('#docs').html("");
            fileCollection = [];
            generateReport();
        });
    });

    $("#downloadReportBtn").click(function(event) {
        event.preventDefault();
        createFile(fileCollection, "GeneratedReport");
    });
});

function genCheckboxItem(title, url, elem) {
    var checkboxItem = $('<input type="checkbox" />').attr({
            id: title,
            'class': 'subsite',
            name: 'subsite',
            value: url
        }),
        label = $('<label for="' + title + '"/>').html(title);
    $(elem).append(checkboxItem, label, $('<br>'));
}

$().SPServices({
    operation: "GetAllSubWebCollection",
    async: true,
    completefunc: function(xData) {
        console.log(xData.responseText);
        $(xData.responseXML).find("Webs > Web").each(function() {
            var $node = $(this);
            console.log($node.attr("Title") + ", " + $node.attr("Url"));
            genCheckboxItem($node.attr("Title"), $node.attr("Url"), '#subsites');
        });
    }
});

function getDocumentInfo() {
    //return the data for current list field that matches the document type
    return function getData(data, error) {
        //show errors in console if exist
        if (error !== undefined) {
            console.log(error);
        }
        console.log("Retrtieving documents for list...");
        //get info for fields returned
        for (var j = 0; j < data.length; j++) {
            //an array to hold the metadata information for each file
            var metaArray = [],
                rsc = data[j].getAttribute("TRIM"),
                fiscalYear = data[j].getAttribute("FY"),
                createdBy = $SP().cleanResult(data[j].getAttribute("Author")),
                absURL = data[j].getAttribute("EncodedAbsUrl"),
                modifiedBy = $SP().cleanResult(data[j].getAttribute("Editor")),
                documentType = $SP().cleanResult(data[j].getAttribute(sName)),
                ctypeID = data[j].getAttribute("ContentTypeId").substring(0, 6);

            //return fields thad only match the "Document" content type, which has id of 0x0101
            if (ctypeID == "0x0101") {
                //log raw document file name and author in the console
                console.log(data[j].getAttribute("FileLeafRef"));
               
                //create new list item element
                var docNode = document.createElement("li"),
                    //save the raw document file name
                    fileName = $SP().cleanResult(data[j].getAttribute("FileLeafRef")),
                    //prepare the string to be used as item in the ordered list #docs
                    listItem = document.createTextNode(
                        fileName +
                        ", URL: " + absURL +
                        ", RS Code: " + rsc +
                        ", Fiscal Year: " + fiscalYear +
                        ", Created By: " + createdBy +
                        ", Modified By: " + modifiedBy
                    );
                //push metadata info into the current scope array
                metaArray.push(fileName, documentType, rsc, fiscalYear, createdBy, modifiedBy, absURL);

                //append the current list itema to the list
                docNode.appendChild(listItem);
                //get the ordered list and append the list item
                document.getElementById("docs").appendChild(docNode);
            }
            //push current file metadata array to the global file collection array
            fileCollection.push(metaArray);
        }
    };
}

function getDocuments(url, recType, staticName) {
    $SP().lists({
        url: url
    }, function(list) {
        for (var i = 0; i < list.length; i++) {
            if (recType !== "All Types") {
                $SP().list(list[i].Name, url).get({
                    fields: "EncodedAbsUrl, Editor,TRIM,FileLeafRef,ContentTypeId,Author,FY" + staticName,
                    where: staticName + '="' + recType + '"'
                }, getDocumentInfo());
            } else {
                $SP().list(list[i].Name, url).get({
                    fields: "EncodedAbsUrl, Editor,TRIM,FileLeafRef,ContentTypeId,Author,FY," + staticName,
                    // where: staticName + '="Active Record" OR ' + staticName + '="Inactive Record" OR ' + staticName + '="Unspecified" OR ' + staticName + '="Non-Record" OR ' + staticName + '=" "'
                }, getDocumentInfo());
            }
        }
    });
}

function generateReport() {
    console.log("Getting documents...");
    var recType = $('#recordTypes option:selected').text();

    for (var i = 0; i < subsites.length; i++) {
        getDocuments(subsites[i], recType, sName);
    }
}
