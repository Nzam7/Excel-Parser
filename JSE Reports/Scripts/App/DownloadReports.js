$(document).ready(function () {

    $("#DownloadReports").off("click").on("click", function () {
        downloadReports();
    });

    $("#UploadReports").off("click").on("click", function () {
        showToaster('info', 'Upload Started', 'Report Upload is underway, you will be notified once it is completed');
        uploadReports();
    });

});

function downloadReports() {
    $.ajax({
        async: false,
        type: "POST",
        url: "DownloadReports.aspx/DownloadAllReports",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            var progressMessage = response.d;
            console.log(progressMessage);
            if (progressMessage === "Success") {
                showToaster('success', 'Completed', 'Report Download Completed'); 
                $("#UploadReports").show();
            }
        },
        failure: function (response) {
            AJAX_Failure(response);
        },
        error: function (xhr) {
            AJAX_Error(xhr);
        }
    }); // end of ajax.   
}

function uploadReports() {


    $.ajax({
        async: false,
        type: "POST",
        url: "DownloadReports.aspx/UploadAllReports",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            var progressMessage = response.d;
            console.log(progressMessage);
            if (progressMessage === "Success") {
                showToaster('success', 'Completed', 'Report Upload Completed');
            }
        },
        failure: function (response) {
            AJAX_Failure(response);
        },
        error: function (xhr) {
            AJAX_Error(xhr);
        }
    }); // end of ajax.   
}

function uploadReportsAsync() {
    $.ajax({
        type: "POST",
        url: "DownloadReports.aspx/UploadAllReportsAsync",
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (response) {
            var progressMessage = response.d;
            console.log(progressMessage);
            if (progressMessage === "Success") {
                showToaster('success', 'Completed', 'Report Upload Completed');
            }
        },
        failure: function (response) {
            AJAX_Failure(response);
        },
        error: function (xhr) {
            AJAX_Error(xhr);
        }
    }); // end of ajax.   
}