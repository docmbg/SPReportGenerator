importScripts("https://rawgit.com/docmbg/SPReportGenerator/1.0/libraries/xlsx.core.min.js");

onmessage = function (epObject) {

    var ep = epObject.data[0];
    var msg = "working";
    postMessage(msg);

    console.log("Began xslx document creation!");

    var wbout = XLSX.write(ep.oFile, { bookType: 'xlsx', bookSST: true, type: 'binary' });

    postMessage(wbout);

    console.log("Done creating xlsx document!");
}