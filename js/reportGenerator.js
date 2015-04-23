/* jshint -W110 */

var subsites = [],
    excelHeader = [
        "Type",
        "Name",
        "Size",
        "Size in bytes",
        "Document Type",
        "FY",
        "Record Series Code",
        "Library",
        "Library Location",
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
    VERSION = "v1.1";

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
    if (rows.length > $("#itemLimit").val()) $("#progress").css("color", "red");
    $("#progress").html(" " + rows.length + " FILES COLLECTED");
    $("#cancelProgress").hide();
    $("#downloadReportBtn").attr("class", "btn-floating btn-large waves-effect waves-light modal-trigger").one("click", function(event) {
        event.preventDefault();
        if (rows.length > $("#itemLimit").val()) {
            $('#warnItemLimit').openModal({
                complete: function() {
                        $("#downloadReportBtn").attr("class", "btn-floating btn-large disabled");
                    } // Callback for Modal close
            });
        } else {
            if (!!window.Worker) {
                var worker = new Worker("js/saveFileWorker.js");
                worker.onmessage = function(e) {
                    if (e.data == "working") {
                        Materialize.toast('Generating excel file. Be patient!', 4000);
                        $("#downloadReportBtn").html("<i class='mdi-action-cached left rotate'></i>").attr("class", "btn-floating btn-large orange").unbind("click");
                    } else {
                        saveExcelFile(e.data[0], $('#recordTypes').find('option:selected').text() + "_" + today);
                        $("#downloadReportBtn").attr("class", "btn-floating btn-large green accent-3").html("<i class='mdi-action-done left'></i>");
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
// function getStaticName() {
//     var currentSite;

//     getCurrentSite().done(function(dfdResolve) {
//         currentSite = dfdResolve;
//     });

//     var dfd = $.Deferred();

//     $SP().list("Inactive Records", currentSite).info(function(fields) {
//         for (var i = 0; i < fields.length; i++) {
//             if (fields[i].DisplayName === "Document Type") {

//                 dfd.resolve(fields[i].StaticName);

//                 for (var choice in fields[i].Choices) {
//                     var recTypeSelect = document.getElementById("recordTypes");
//                     recTypeSelect.options[recTypeSelect.options.length] =
//                         new Option(fields[i].Choices[choice], fields[i].Choices[choice]);
//                     $('select').material_select();
//                 }
//             }
//         }
//     });
//     return dfd.promise();
// }

// getStaticName().done(function(a) {
//     sName = a;
// });

function genSitesArray() {
    console.log("Generating array of sub-sites");
    subsites = $("#subsites").find("input:checkbox:checked").map(function() {
        return $(this).val();
    }).get();
    return $.Deferred().resolve(false);
}


function resetValues() {
    $("#progress").css("color", "#0096D6");
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
                Materialize.toast('Please select at least one site/sub-site!', 2000);
                getFiles();
            } else {
                generateReport();
            }
        });
    });
};

$(document).ready(function() {
    $('select').material_select();
    $('.modal-trigger').leanModal();
    $("#instrContent").load("https://rawgit.com/docmbg/SPReportGenerator/1.0/helpers/instructions.html");
    $("#changelogContent").load("https://rawgit.com/docmbg/SPReportGenerator/1.0/helpers/changelog.html");
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

//catches window error
function catchError() {
    window.onerror = function(errorMsg) {
        //alert('Error: ' + errorMsg + ' Script: ' + url + ' Line: ' + lineNumber);
        procdLists.push(errorMsg);
    };
}

function bytesToSize(bytes) {
    if (bytes === 0) return '0 Byte';
    var k = 1000;
    var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
    var i = Math.floor(Math.log(bytes) / Math.log(k));
    return (bytes / Math.pow(k, i)).toPrecision(3) + ' ' + sizes[i];
}

function getDocumentInfo(listName, siteURL) {
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
                    fileSize,
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
                fileSize = $SP().cleanResult(data[j].getAttribute("File_x0020_Size"));
                //documentType = $SP().cleanResult(data[j].getAttribute(sName)).toString();
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

                    if (fileSize !== null) {
                        ep.write({
                            "cell": "C" + (rows.length + 1),
                            "content": bytesToSize(fileSize)
                        });
                    }

                    if (fileSize !== null) {
                        ep.write({
                            "cell": "D" + (rows.length + 1),
                            "content": fileSize
                        });
                    }

                    // if (documentType.length >= 1) {
                    //     ep.write({
                    //         "cell": "E" + (rows.length + 1),
                    //         "content": documentType
                    //     });
                    // }

                    if (fiscalYear !== null) {
                        ep.write({
                            "cell": "F" + (rows.length + 1),
                            "content": fiscalYear
                        });
                    }

                    if (recordCode !== null) {
                        ep.write({
                            "cell": "G" + (rows.length + 1),
                            "content": recordCode
                        });
                    }

                    ep.write({
                        "cell": "H" + (rows.length + 1),
                        "content": listName + ""
                    });

                    ep.write({
                        "cell": "I" + (rows.length + 1),
                        "content": siteURL.split("teams/")[1].replace(/\//g, " > ") + ""
                    });

                    if (createdBy.length >= 1) {
                        ep.write({
                            "cell": "J" + (rows.length + 1),
                            "content": createdBy + ""
                        });
                    }

                    if (modifiedBy.length >= 1) {
                        ep.write({
                            "cell": "K" + (rows.length + 1),
                            "content": modifiedBy + ""
                        });
                    }

                    if (created !== null) {
                        ep.write({
                            "cell": "L" + (rows.length + 1),
                            "content": created + ""
                        });
                    }

                    if (modified !== null) {
                        ep.write({
                            "cell": "M" + (rows.length + 1),
                            "content": modified + ""
                        });
                    }

                    if (absURL !== null) {
                        ep.write({
                            "cell": "N" + (rows.length + 1),
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
            $("#downloadReportBtn").attr("class", "btn-floating btn-large disabled").unbind("click");
            $("#cancelProgress").show().click(function(e) {
                e.preventDefault();
            });
        } else {
            executeFileSave();
        }
    };
}

function getDocuments(url, recType, staticName) {

    $SP().lists({
        url: url
    }, function(list) {

        for (var i = 0; i < list.length; i++) {

            listCount.push(list[i].Name);
            if (recType !== "All Types") {
                $SP().list(list[i].Name, url).get({
                    fields: "ContentType,ContentTypeId,File_x0020_Size,DocIcon,FileLeafRef,FY,TRIM,EncodedAbsUrl,Editor,Author,Created_x0020_Date,Last_x0020_Modified," + staticName,
                    where: staticName + '="' + recType + '"' + ' AND ContentType = "Document"',
                    expandUserField: true
                }, getDocumentInfo(list[i].Name, url));
            } else {
                $SP().list(list[i].Name, url).get({
                    fields: "ContentType,ContentTypeId,File_x0020_Size,DocIcon,FileLeafRef,FY,TRIM,EncodedAbsUrl,Editor,Author,Created_x0020_Date,Last_x0020_Modified,",
                    where: 'ContentType = "' + 'Document"',
                    expandUserField: true
                        //staticName + '="Active Record" OR ' + staticName + '="Inactive Record" OR ' + staticName + '="Unspecified" OR ' + staticName + '="Non-Record" OR ' + staticName + '=" "'
                }, getDocumentInfo(list[i].Name, url));
            }
        }
    });
}

function generateReport() {
    var recType;
    console.log("Getting documents...");
    recType = $('#recordTypes').find('option:selected').val();

    for (var i = 0; i < subsites.length; i++) {
        getDocuments(subsites[i], recType, sName);
    }
}
