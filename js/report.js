/* jshint -W110 */

var subsites = [],
    excelHeader = [
        "Type",
        "Name",
        "Document Type",
        "FY",
        "Record Series Code",
        "Created By",
        "Modified By",
        "Created",
        "Last Modified",
        "URL"
    ],
    ep = new ExcelPlus(),
    today = new Date(),
    dd = today.getDate(),
    mm = today.getMonth() + 1,
    year = today.getFullYear(),
    today = year + "-" + mm + "-" + dd,
    sName,
    listCount = [],
    procdLists = [],
    rows = [],
    rowNumber = 0,
    activeRecs = [],
    inactiveRecs = [],
    unspecRecs = [],
    nonRecs = [],
    otherDocs = [],
    VERSION = "ver. 0.1b";

function getPercent(num, total) {
    var result = 0;
    if (rows.length > 0) {
        var percent = (num / total) * 100;
        result = parseFloat(Math.round(percent * 10) / 10).toFixed(1).replace(/\.0/, "");
    }
    return result;
}

function executeFileSave() {
    $.p.end();
    getFiles();
    $("#getFilesBtn").attr("class", "waves-effect waves-light btn");
    $("#activeRecs").html(activeRecs.length + " Files <br>" + getPercent(activeRecs.length, rows.length) + "% of Total");
    $("#inactiveRecs").html(inactiveRecs.length + " Files <br>" + getPercent(inactiveRecs.length, rows.length) + "% of Total");
    $("#unspecRecs").html(unspecRecs.length + " Files <br>" + getPercent(unspecRecs.length, rows.length) + "% of Total");
    $("#nonRecs").html(nonRecs.length + " Files <br>" + getPercent(nonRecs.length, rows.length) + "% of Total");
    $("#otherDocs").html(otherDocs.length + " Files <br>" + getPercent(otherDocs.length, rows.length) + "% of Total");
    $("#progress").html(" " + rows.length + " FILES COLLECTED");
    $("#downloadReportBtn").attr("class", "btn-floating btn-large waves-effect waves-light modal-trigger");
    $("#cancelProgress").hide();
    $("#downloadReportBtn").one("click", function(event) {
        event.preventDefault();
        //TODO check for number of items and limit it to 25000
        if (rows.length > 25000) {
            $('#warnItemLimit').openModal({
            	complete: function() {
            		$("#downloadReportBtn").attr("class", "btn-floating btn-large disabled");
            	} // Callback for Modal close
            });
        } else {
            if (!!window.Worker) {
                var worker = new Worker("../js/saveFileWorker.js");
                worker.onmessage = function(e) {
                    if (e.data == "working") {
                        Materialize.toast('Generating excel file. Be patient!', 4000);
                        $("#downloadReportBtn").html("<i class='mdi-action-cached left rotate'></i>");
                        $("#downloadReportBtn").attr("class", "btn-floating btn-large orange");
                        $("#downloadReportBtn").unbind("click");
                    } else {
                        saveExcelFile(e.data, $('#recordTypes option:selected').text() + "_" + today);
                        $("#downloadReportBtn").attr("class", "btn-floating btn-large green accent-3");
                        $("#downloadReportBtn").html("<i class='mdi-action-done left'></i>");
                    }
                };
                worker.postMessage([ep]);
            }
        }
    });
}

function saveExcelFile(data, fileName) {
    //set the file name
    var filename = fileName + ".xlsx";

    //put the file stream together
    var s2ab = function(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    };
    //invoke the saveAs method from FileSaver.js
    saveAs(new Blob([s2ab(data)], {
        type: "application/octet-stream"
    }), filename);
}

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
    subsites = $("#subsites").find("input:checkbox:checked").map(function() {
        return $(this).val();
    }).get();
    return $.Deferred().resolve(false);
}


function resetValues() {
    $('#reportContainer').find('>ul>li>p').html("0 Files <br>0% of Total");
    ep.createFile("Report");
    ep.write({
        "content": [excelHeader]
    });
    subsites = [];
    activeRecs = [];
    inactiveRecs = [];
    unspecRecs = [];
    nonRecs = [];
    rows = [];
    listCount = [];
    procdLists = [];
    otherDocs = [];
    $("#downloadReportBtn").html("<i class='mdi-file-file-download left'></i>");
}


var getFiles = function() {
    $("#getFilesBtn").one("click", function(event) {
        resetValues();
        $.p = progressJs("#progressBar").start();
        event.preventDefault();
        genSitesArray().done(function() {
            if (subsites.length === 0) {
                Materialize.toast('Please select at least one stie/subsite!', 2000);
                getFiles();
            } else {
                generateReport();
            }
        });
    });
};

$(document).ready(function() {
    $('.modal-trigger').leanModal();
    $("#instrContent").load("https://rawgit.com/docmbg/SPReportGenerator/beta/helpers/instructions.html");
    $("#changelogContent").load("https://rawgit.com/docmbg/SPReportGenerator/beta/helpers/changelog.html");
    $("#version").find(">a").html(VERSION);
    $(".button-collapse").sideNav();
    $("#cancelProgress").hide();
    $("#checkAll").change(function() {
        $("input:checkbox").prop('checked', $(this).prop("checked"));
    });
    getFiles();
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
        $(xData.responseXML).find("Webs > Web").each(function() {
            var $node = $(this);
            genCheckboxItem($node.attr("Title"), $node.attr("Url"), '#subsites');
        });
    }
});


function progress(count, all) {
    var result = Math.floor((count.length / all.length) * 100);

    var textNode = document.createTextNode(" " + result + "% OF QUERY FILES COLLECTED");
    $("#progress").html(textNode);

    return result;
}

function catchError() {
    window.onerror = function(errorMsg) {
        //alert('Error: ' + errorMsg + ' Script: ' + url + ' Line: ' + lineNumber);
        procdLists.push(errorMsg);
    };
}

function getDocumentInfo() {
    //return the data for current list field that matches the document type
    return function(data, error) {
        //show errors in console if exist
        if (!!error) {
            console.log(error);
        }

        var regExEmail = new RegExp(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/i);
        //console.log("Retrtieving documents for list...");
        //get info for fields returned
        for (var j = 0; j < data.length; j++) {
            //ignore sharepointplus ajax error
            //an array to hold the metadata information for each
            catchError();

            var cTypeID = data[j].getAttribute("ContentTypeId").substring(0, 6);

            if (cTypeID === "0x0101") {
                var type,
                    fileName,
                    documentType,
                    fiscalYear,
                    recordCode,
                    createdBy,
                    modifiedBy,
                    created,
                    modified,
                    absURL;


                type = data[j].getAttribute("DocIcon");
                fileName = $SP().cleanResult(data[j].getAttribute("FileLeafRef"));
                documentType = $SP().cleanResult(data[j].getAttribute(sName)).toString();
                fiscalYear = data[j].getAttribute("FY");
                recordCode = data[j].getAttribute("TRIM");
                createdBy = $SP().cleanResult(data[j].getAttribute("Author")
                    .match(regExEmail)).toString().split(',')[0];
                modifiedBy = $SP().cleanResult(data[j].getAttribute("Editor")
                    .match(regExEmail)).toString().split(',')[0];
                created = $SP().cleanResult(data[j].getAttribute("Created_x0020_Date"));
                modified = $SP().cleanResult(data[j].getAttribute("Last_x0020_Modified"));
                absURL = data[j].getAttribute("EncodedAbsUrl");

                //return fields thad only match the "Document" content type, which has id of 0x0101
                if (
                    type !== "master" &&
                    type !== "aspx" &&
                    type !== "png" &&
                    type !== "gif" &&
                    type !== "xsl" &&
                    type !== "xoml" &&
                    type !== "jpg" &&
                    type !== "xsn" &&
                    type !== "css" &&
                    type !== "xaml" &&
                    type !== "rules"
                ) {

                    rowNumber += 1;
                    rows.push(rowNumber);

                    switch (documentType) {
                        case 'Active Record':
                            activeRecs.push('a');
                            break;
                        case 'Inactive Record':
                            inactiveRecs.push('i');
                            break;
                        case 'Non-Record':
                            nonRecs.push('n');
                            break;
                        case 'Unspecified':
                            unspecRecs.push('u');
                            break;
                        default:
                            otherDocs.push('o');
                    }

                    if (type !== null) {
                        ep.write({
                            "cell": "A" + (rows.length + 1),
                            "content": type
                        });
                    }

                    if (fileName !== null) {
                        ep.write({
                            "cell": "B" + (rows.length + 1),
                            "content": fileName
                        });
                    }

                    if (documentType.length >= 1) {
                        ep.write({
                            "cell": "C" + (rows.length + 1),
                            "content": documentType
                        });
                    }

                    if (fiscalYear !== null) {
                        ep.write({
                            "cell": "D" + (rows.length + 1),
                            "content": fiscalYear
                        });
                    }

                    if (recordCode !== null) {
                        ep.write({
                            "cell": "E" + (rows.length + 1),
                            "content": recordCode
                        });
                    }

                    if (createdBy.length >= 1) {
                        ep.write({
                            "cell": "F" + (rows.length + 1),
                            "content": createdBy + ""
                        });
                    }

                    if (modifiedBy.length >= 1) {
                        ep.write({
                            "cell": "G" + (rows.length + 1),
                            "content": modifiedBy + ""
                        });
                    }

                    if (created !== null) {
                        ep.write({
                            "cell": "H" + (rows.length + 1),
                            "content": created + ""
                        });
                    }

                    if (modified !== null) {
                        ep.write({
                            "cell": "I" + (rows.length + 1),
                            "content": modified + ""
                        });
                    }

                    if (absURL !== null) {
                        ep.write({
                            "cell": "J" + (rows.length + 1),
                            "content": absURL + ""
                        });
                    }

                }

                //create new list item element
                //var li = "<li class='collection-item avatar'><i class='mdi-file-folder circle blue'></i>" +
                //    "<span class='title'><b>" + fileName + "</b></span>" +
                //    "<p><b>Created by: </b>" + createdBy + "<br>" +
                //    "<b>Modified by: </b>" + modifiedBy + "</p>" +
                //    "<a href='" + absURL + "' class='secondary-content'><i class='mdi-file-file-download'></i></a>" +
                //    "</li>";

                //get the ordered list and append the list item, doesn't currently exist in html
                //$("#docs").append(li);
            }

            //TODO
            //push li variable to the jStorage local storage when created
            //it will allow the users to browse the files
        }
        procdLists.push('d');
        $.p.set(progress(procdLists, listCount));
        if (progress(procdLists, listCount) < 100) {
            $("#getFilesBtn").attr("class", "btn disabled");
            $("#downloadReportBtn").attr("class", "btn-floating btn-large disabled");
            $("#downloadReportBtn").unbind("click");
            $("#cancelProgress").show();
            $("#cancelProgress").click(function(e) {
                e.preventDefault();
            });
        } else {
            executeFileSave();
        }
    };
}

// function getDocInfo() {

//     return function() {
//         //$(this.data.responseXML).SPFilterNode("z:row").SPXmlToJson();
//         //assign json data object for the list to the data variable
//         var data = this.data;
//         //find email in address book contact
//         //var regExEmail = new RegExp(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/i);
//         //iterate the data
//         for (var i = 0; i < data.length; i++) {

//             catchError();
//             var cTypeID;
//             if (!!data[i].ContentTypeId) cTypeID = data[i].ContentTypeId.substring(0, 6);

//             if (cTypeID === "0x0101") {
//                 //declare vars to hold different metadata
//                 var type,
//                     fileName,
//                     documentType,
//                     fiscalYear,
//                     recordCode,
//                     createdBy,
//                     modifiedBy,
//                     created,
//                     modified,
//                     absURL;

//                 //assign metadata accordingly
//                 if (!!data[i].File_x0020_Type) type = data[i].File_x0020_Type;

//                 if (!!data[i].FileLeafRef) fileName = $SP().cleanResult(data[i].FileLeafRef);

//                 if (!!data[i][sName]) documentType = data[i][sName];

//                 if (!!data[i].FY) fiscalYear = data[i].FY;

//                 if (!!data[i].TRIM) recordCode = data[i].TRIM;

//                 if (!!data[i].Editor.email) modifiedBy = data[i].Editor.email;

//                 if (!!data[i].Author.email) createdBy = data[i].Author.email;

//                 if (!!data[i].Created_x0020_Date.lookupValue) created = data[i].Created_x0020_Date.lookupValue.substring(0, 10);

//                 if (data[i].Last_x0020_Modified.lookupValue) modified = data[i].Last_x0020_Modified.lookupValue.substring(0, 10);

//                 if (data[i].FileRef.lookupValue) absURL = data[i].FileRef.lookupValue;

//                 //log a preview of the file

//                 // console.log("Type: " + type + "\n" +
//                 //     "Name: " + fileName + "\n" +
//                 //     "Document Type: " + documentType + "\n" +
//                 //     "Fiscal Year: " + fiscalYear + "\n" +
//                 //     "RSC: " + recordCode + "\n" +
//                 //     "Modified By: " + modifiedBy + "\n" +
//                 //     "Created By: " + createdBy + "\n" +
//                 //     "Created: " + created + "\n" +
//                 //     "Modified: " + modified + "\n" +
//                 //     "URL: " + absURL
//                 // );

//                 //return fields thad only match the "Document" content type, which has id of 0x0101
//                 if (
//                     type !== "master" &&
//                     type !== "aspx" &&
//                     type !== "png" &&
//                     type !== "gif" &&
//                     type !== "xsl" &&
//                     type !== "xoml" &&
//                     type !== "jpg" &&
//                     type !== "xsn" &&
//                     type !== "css" &&
//                     type !== "xaml" &&
//                     type !== "rules"
//                 ) {

//                     rowNumber += 1;
//                     rows.push(rowNumber);

//                     switch (documentType) {
//                         case 'Active Record':
//                             activeRecs.push('a');
//                             break;
//                         case 'Inactive Record':
//                             inactiveRecs.push('i');
//                             break;
//                         case 'Non-Record':
//                             nonRecs.push('n');
//                             break;
//                         case 'Unspecified':
//                             unspecRecs.push('u');
//                             break;
//                         default:
//                             otherDocs.push('o');
//                     }

//                     if (type !== undefined) {
//                         ep.write({
//                             "cell": "A" + (rows.length + 1),
//                             "content": type
//                         });
//                     }

//                     if (fileName !== undefined) {
//                         ep.write({
//                             "cell": "B" + (rows.length + 1),
//                             "content": fileName
//                         });
//                     }

//                     if (documentType !== undefined) {
//                         ep.write({
//                             "cell": "C" + (rows.length + 1),
//                             "content": documentType
//                         });
//                     }

//                     if (fiscalYear !== undefined) {
//                         ep.write({
//                             "cell": "D" + (rows.length + 1),
//                             "content": fiscalYear
//                         });
//                     }

//                     if (recordCode !== undefined) {
//                         ep.write({
//                             "cell": "E" + (rows.length + 1),
//                             "content": recordCode
//                         });
//                     }

//                     if (createdBy !== undefined) {
//                         ep.write({
//                             "cell": "F" + (rows.length + 1),
//                             "content": createdBy + ""
//                         });
//                     }

//                     if (modifiedBy !== undefined) {
//                         ep.write({
//                             "cell": "G" + (rows.length + 1),
//                             "content": modifiedBy + ""
//                         });
//                     }

//                     if (created !== undefined) {
//                         ep.write({
//                             "cell": "H" + (rows.length + 1),
//                             "content": created + ""
//                         });
//                     }

//                     if (modified !== undefined) {
//                         ep.write({
//                             "cell": "I" + (rows.length + 1),
//                             "content": modified + ""
//                         });
//                     }

//                     if (absURL !== undefined) {
//                         ep.write({
//                             "cell": "J" + (rows.length + 1),
//                             "content": absURL + ""
//                         });
//                     }

//                 }
//                 //create new list item element
//                 //var li = "<li class='collection-item avatar'><i class='mdi-file-folder circle blue'></i>" +
//                 //    "<span class='title'><b>" + fileName + "</b></span>" +
//                 //    "<p><b>Created by: </b>" + createdBy + "<br>" +
//                 //    "<b>Modified by: </b>" + modifiedBy + "</p>" +
//                 //    "<a href='" + absURL + "' class='secondary-content'><i class='mdi-file-file-download'></i></a>" +
//                 //    "</li>";

//                 //get the ordered list and append the list item, doesn't currently exist in html
//                 //$("#docs").append(li);
//             }
//             //TODO
//             //push li variable to the jStorage local storage when created
//             //it will allow the users to browse the files
//         }
//         console.log("==============finished for library================");
//         //procdLists.push('d');
//         $.p.set(progress(procdLists, listCount));
//         if (progress(procdLists, listCount) < 100) {
//             $("#getFilesBtn").attr("class", "btn disabled");
//             $("#downloadReportBtn").attr("class", "btn-floating btn-large disabled");
//             $("#downloadReportBtn").unbind("click");
//             $("#cancelProgress").show();
//             $("#cancelProgress").click(function(e) {
//                 e.preventDefault();
//             });
//         } else {
//             executeFileSave();
//         }
//     };
// }

function getDocuments(url, recType, staticName) {

    // var camlViewFields = "<ViewFields>" +
    //     "<FieldRef Name='File_x0020_Type'/>" +
    //     "<FieldRef Name='ContentTypeId'/>" +
    //     "<FieldRef Name='FileExtension'/>" +
    //     "<FieldRef Name='FileLeafRef'/>" +
    //     "<FieldRef Name='" + staticName + "'/>" +
    //     "<FieldRef Name='FY'/>" +
    //     "<FieldRef Name='TRIM'/>" +
    //     "<FieldRef Name='Editor'/>" +
    //     "<FieldRef Name='Author'/>" +
    //     "<FieldRef Name='Created_x0020_Date'/>" +
    //     "<FieldRef Name='Last_x0020_Modified'/>" +
    //     "<FieldRef Name='FileRef'/>" +
    //     "</ViewFields>";

    $SP().lists({
        url: url
    }, function(list) {

        for (var i = 0; i < list.length; i++) {

            // $().SPServices({
            //     webURL: url,
            //     operation: "GetListItems",
            //     async: true,
            //     listName: list[i].Name,
            //     CAMLViewFields: camlViewFields,
            //     completefunc: function(xData, Status) {
            //         var myJson = $(xData.responseXML).SPFilterNode("z:row").SPXmlToJson({
            //             mapping: {
            //                 ows_FileLeafRef: {
            //                     mappedName: "Name",
            //                     objectType: "Text"
            //                 },
            //                 ows_Created_x0020_Date: {
            //                     mappedName: "Created",
            //                     objectType: "Text"
            //                 },
            //                 ows_Author: {
            //                     mappedName: "Author",
            //                     objectType: "User"
            //                 },
            //                 ows_ContentTypeId: {
            //                     mappedName: "ctypeID",
            //                     objectType: "Text"
            //                 },
            //                 ows_File_x0020_Type: {
            //                     mappedName: "FileExtension",
            //                     objectType: "Text"
            //                 },
            //                 ows_EncodedAbsUrl: {
            //                     mappedName: "URL",
            //                     objectType: "Text"
            //                 }
            //             }, // name, mappedName, objectType
            //             includeAllAttrs: false
            //         });
            //     console.log(myJson);
            //     procdLists.push('d');
            //     }
            // });

            // var listPromise = $().SPServices.SPGetListItemsJson({
            //     webURL: url,
            //     listName: list[i].Name,
            //     CAMLViewFields: camlViewFields,
            //     CAMLQueryOptions: "<ExpandUserField>True</ExpandUserField>",
            //     // mapping: null,
            //     // mappingOverrides: null,
            //     debug: true
            // });

            // $.when(listPromise).then(getDocInfo());

            listCount.push(list[i].Name);
            if (recType !== "All Types") {
                $SP().list(list[i].Name, url).get({
                    fields: "ContentType,ContentTypeId,DocIcon,FileLeafRef,FY,TRIM,EncodedAbsUrl,Editor,Author,Created_x0020_Date,Last_x0020_Modified," + staticName,
                    where: staticName + '="' + recType + '"' + ' AND ContentType = "Document"',
                    expandUserField: true
                }, getDocumentInfo());
            } else {
                $SP().list(list[i].Name, url).get({
                    fields: "ContentType,ContentTypeId,DocIcon,FileLeafRef,FY,TRIM,EncodedAbsUrl,Editor,Author,Created_x0020_Date,Last_x0020_Modified," + staticName,
                    where: 'ContentType = "' + 'Document"',
                    expandUserField: true
                        //staticName + '="Active Record" OR ' + staticName + '="Inactive Record" OR ' + staticName + '="Unspecified" OR ' + staticName + '="Non-Record" OR ' + staticName + '=" "'
                }, getDocumentInfo());
            }
        }
    });
}

function generateReport() {
    console.log("Getting documents...");
    var recType = $('#recordTypes').find('option:selected').val();

    for (var i = 0; i < subsites.length; i++) {
        getDocuments(subsites[i], recType, sName);
    }
}
