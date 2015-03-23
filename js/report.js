/* jshint -W110 */

var subsites = [];


//variable to hold the static name of the Document Type column
var sName;

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
    }).get(); // <----
    console.log(subsites);
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
            generateReport();
        });
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

        //get info for fields returned
        for (var j = 0; j < data.length; j++) {
            var trim = data[j].getAttribute("TRIM");
            var fiscalYear = data[j].getAttribute("FY");
            var createdBy = data[j].getAttribute("Created_x0020_By");

            //get document content type guid
            var ctypeID = data[j].getAttribute("ContentTypeId").substring(0, 6);

            //return fields thad only match the "Document" content type, which is 0x0101
            if (ctypeID == "0x0101") {
                //log raw document file name and author in the console
                console.log(data[j].getAttribute("FileLeafRef") +
                    ", " + data[j].getAttribute("Created_x0020_By"));

                //save the raw document file name
                var rawname = $SP().cleanResult(data[j].getAttribute("FileLeafRef"));

                //clean up the document file name and save it in a new variable to be used in the excel table
                var docname = document.createTextNode(rawname +
                    ", " + "RS Code:" + trim + ", " + fiscalYear + ", " + createdBy);

                //create new list item element
                var docNode = document.createElement("li");

                //append the current document name to the list item
                docNode.appendChild(docname);

                //get the main div and append the 
                document.getElementById("docs").appendChild(docNode);
            }
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
                    fields: "TRIM,FileLeafRef,ContentTypeId,Created_x0020_By,FY",
                    where: staticName + '="' + recType + '"'
                        //static way
                        //'Document_x0020_Type = "' + recType + '"'
                }, getDocumentInfo());
            } else {
                $SP().list(list[i].Name, url).get({
                    fields: "TRIM,FileLeafRef,ContentTypeId,Created_x0020_By,FY",
                }, getDocumentInfo());
            }
        }
    });
}

function generateReport() {
    console.log("Getting document...");
    var subs = subsites;
    var recType = $('#recordTypes option:selected').text();
    //depricated
    // var staticName = $SP().toXSLString($('#documentType option:selected').val());
    for (var i = 0; i < subs.length; i++) {
        getDocuments(subs[i], recType, sName);
    }
}

//simo added something 2
//mater branch comment