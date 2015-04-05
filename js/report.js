/* jshint -W110 */

var subsites = [],
    fileCollection = [excelHeader],
    excelHeader = ["Type", "Name", "Document Type", "FY", "Record Series Code", "Created By", "Modified By", "Created", "Last Modifed", "URL"],
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
    rowNumber = 0;

function executeFileSave() {
    $.p.end();
    $("#progress").html(" " + rows.length + " FILES COLLECTED");
    $("#downloadReportBtn").attr("class", "btn-floating btn-large");
    $("#cancelProgress").hide();
    $("#downloadReportBtn").click(function (event) {
        event.preventDefault();
        if (!!window.Worker) {
            var worker = new Worker("../js/saveFileWorker.js");
            worker.onmessage = function (e) {
                if (e.data == "working") {
                    Materialize.toast('Generating excel file. Be patient!', 4000) // 2000 is the duration of the toast
                    $("#downloadReportBtn").unbind("click");
                    $("#downloadReportBtn").html("<i class='mdi-action-cached left'></i>");
                    $("#downloadReportBtn").attr("class", "btn-floating btn-large disabled");
                } else {
                    saveExcelFile(e.data, $('#recordTypes option:selected').text() + "_" + today);
                    $("#downloadReportBtn").attr("class", "btn-floating btn-large green accent-3");
                    $("#downloadReportBtn").html("<i class='mdi-action-done left'></i>");
                }
            }
            worker.postMessage([ep]);
        }
    });
}

function saveExcelFile(data, fileName) {
    //set the file name
    var filename = fileName + ".xlsx";
    //check for browser compatability
    //if (typeof Uint8Array === "undefined") {
    //    this.error = "[saveAs] Sorry but this function is only supported by modern browsers";
    //    this._showErrors();
    //    return this;
    //}
    //put the file stream together
    var s2ab = function (s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    //invoke the saveAs method from FileSaver.js
    saveAs(new Blob([s2ab(data)], { type: "application/octet-stream" }), filename);
}

function getCurrentSite() {
    var dfd = $.Deferred();
    dfd.resolve($().SPServices.SPGetCurrentSite());
    return dfd.promise();
}

// will use in RIST Page
function getStaticName() {
    var currentSite;

    getCurrentSite().done(function (dfdResolve) {
        currentSite = dfdResolve;
    });

    var dfd = $.Deferred();

    $SP().list("Inactive Records", currentSite).info(function (fields) {
        for (var i = 0; i < fields.length; i++) {
            if (fields[i].DisplayName === "Document Type") {

                dfd.resolve(fields[i].StaticName);

                for (var choice in fields[i].Choices) {
                    // console.log(fields[i].Choices[choice]);
                    var recTypeSelect = document.getElementById("recordTypes");
                    recTypeSelect.options[recTypeSelect.options.length] =
                        new Option(fields[i].Choices[choice], fields[i].Choices[choice]);
                }
            }
        }
    });
    return dfd.promise();
}

getStaticName().done(function (a) {
    sName = a;
});

function genSitesArray() {
    console.log("Generating array of sub-sites");
    subsites = $("#subsites input:checkbox:checked").map(function () {
        return $(this).val();
    }).get();
    return $.Deferred().resolve(false);
}

$(document).ready(function () {
    $(".button-collapse").sideNav();
    $("#cancelProgress").hide();
    $("#checkAll").change(function () {
        $("input:checkbox").prop('checked', $(this).prop("checked"));
    });

    $("#getFilesBtn").click(function (event) {
        ep.createFile("Report");
        ep.write({
            "content": [excelHeader]
        });
        $("#downloadReportBtn").html("<i class='mdi-file-file-download left'></i>");
        //$("#docsContainer").css('height', $("#sitesContainer").css('height'));
        $.p = progressJs("#progressBar").start();
        event.preventDefault();
        genSitesArray().done(function () {
            if (subsites.length === 0) {
                $.p.end();
                Materialize.toast('Please select at least one stie/subsite!', 2000) // 2000 is the duration of the toast
            }
            $('#reportCollection >li >p').html("0 Files <br>0% of Total");
            rows = [];
            fileCollection = [excelHeader];
            listCount = [];
            procdLists = [];
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
    completefunc: function (xData) {
        // console.log(xData.responseText);
        $(xData.responseXML).find("Webs > Web").each(function () {
            var $node = $(this);
            // console.log($node.attr("Title") + ", " + $node.attr("Url"));
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

function getDocumentInfo() {

    //return the data for current list field that matches the document type
    return function (data) {
        //show errors in console if exist
        //if (error !== undefined) {
        //    console.log(error);
        //}
        var regExEmail = new RegExp(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/i);
        //console.log("Retrtieving documents for list...");
        //get info for fields returned
        for (var j = 0; j < data.length; j++) {
            //ignore sharepointplus ajax error
            //an array to hold the metadata information for each file

            var ctypeID = data[j].getAttribute("ContentTypeId").substring(0, 6),
                        type = data[j].getAttribute("DocIcon"),
                        fileName = $SP().cleanResult(data[j].getAttribute("FileLeafRef")),
                        documentType = $SP().cleanResult(data[j].getAttribute(sName)).toString(),
                        fiscalYear = data[j].getAttribute("FY"),
                        recordCode = data[j].getAttribute("TRIM"),
                        createdBy = $SP().cleanResult(data[j].getAttribute("Author")
                            .match(regExEmail)).toString().split(',')[0],
                        modifiedBy = $SP().cleanResult(data[j].getAttribute("Editor")
                            .match(regExEmail)).toString().split(',')[0],
                        created = $SP().cleanResult(data[j].getAttribute("Created_x0020_Date")),
                        modified = $SP().cleanResult(data[j].getAttribute("Last_x0020_Modified")),
                        absURL = data[j].getAttribute("EncodedAbsUrl");


            //console.log("Record Code: " + data[j].getAttribute("TRIM") + ";");
            //console.log("Document Type: " + documentType + ";");
            //console.log("Fiscal Year: " + fiscalYear + ";");
            //console.log("Type: " + type);
            //return fields thad only match the "Document" content type, which has id of 0x0101
            if (ctypeID == "0x0101" &&
                type !== "master" &&
                type !== "aspx" &&
                type !== "png" &&
                type !== "gif" &&
                type !== "xsl" &&
                type !== "xoml" &&
                type !== "jpg" &&
                type !== "xml" &&
                type !== "xsn" &&
                type !== "css" &&
                type !== "xaml" &&
                type !== "rules") {

                rowNumber += 1;
                rows.push(rowNumber);

                if (type !== null) {
                    ep.write({
                        "cell": "A" + (rows.length + 1),
                        "content": type
                    });
                }

                if (type !== null) {
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

                //log raw document file name and author in the console
                //console.log(data[j].getAttribute("FileLeafRef"));


                //create new list item element
                //var li = "<li class='collection-item avatar'><i class='mdi-file-folder circle blue'></i>" +
                //    "<span class='title'><b>" + fileName + "</b></span>" +
                //    "<p><b>Created by: </b>" + createdBy + "<br>" +
                //    "<b>Modified by: </b>" + modifiedBy + "</p>" +
                //    "<a href='" + absURL + "' class='secondary-content'><i class='mdi-file-file-download'></i></a>" +
                //    "</li>";

                //get the ordered list and append the list item
                //$("#docs").append(li);
            }
            //push li variable to the jStorage local storage
        }
        procdLists.push('d');
        //progress(procdLists, listCount);
        $.p.set(progress(procdLists, listCount));
        if (progress(procdLists, listCount) < 100) {
            $("#downloadReportBtn").attr("class", "btn-floating btn-large disabled");
            $("#downloadReportBtn").unbind("click");
            $("#cancelProgress").show();
            $("#cancelProgress").click(function () {
                executeFileSave();
                prevent.default();
            });
        } else {
            executeFileSave();
        }
    };
}

function getDocuments(url, recType, staticName) {
    $SP().lists({
        url: url
    }, function (list) {
        for (var i = 0; i < list.length; i++) {
            if (list.length) {
                listCount.push(list[i].Name);
            }
            if (recType !== "All Types") {
                $SP().list(list[i].Name, url).get({
                    fields: "ContentTypeId,DocIcon,FileLeafRef,FY,TRIM,EncodedAbsUrl,Editor,Author,Created_x0020_Date,Last_x0020_Modified," + staticName,
                    where: staticName + '="' + recType + '"',
                    expandUserField: true
                }, getDocumentInfo());
            } else {
                $SP().list(list[i].Name, url).get({
                    fields: "ContentTypeId,DocIcon,FileLeafRef,FY,TRIM,EncodedAbsUrl,Editor,Author,Created_x0020_Date,Last_x0020_Modified," + staticName,
                    expandUserField: true
                    // where: staticName + '="Active Record" OR ' + staticName + '="Inactive Record" OR ' + staticName + '="Unspecified" OR ' + staticName + '="Non-Record" OR ' + staticName + '=" "'
                }, getDocumentInfo());
            }
        }
    });
}

function generateReport() {
    console.log("Getting documents...");
    var recType = $('#recordTypes option:selected').val();

    for (var i = 0; i < subsites.length; i++) {
        getDocuments(subsites[i], recType, sName);
    }
}