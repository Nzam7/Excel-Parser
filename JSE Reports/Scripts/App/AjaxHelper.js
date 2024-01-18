function AJAX_Error(xhr) {
    console.log('Request Status: ' + xhr.status + ' Status Text: ' + xhr.statusText + '  Response Text:' + xhr.responseText);
    var message;
    var statusErrorMap = {
        '400': "Server understood the request, but request content was invalid.",
        '401': "Unauthorized access.",
        '403': "Forbidden resource can't be accessed.",
        '500': "Internal server error.",
        '503': "Service unavailable."
    };
    if (xhr.status) {
        message = statusErrorMap[xhr.status];
        if (!message) {
            message = "Unknown Error \n.";
        }
    } else if (exception === 'parsererror') {
        message = "Error.\nParsing JSON Request failed.";
    } else if (exception === 'timeout') {
        message = "Request Time out.";
    } else if (exception === 'abort') {
        message = "Request was aborted by the server";
    } else {
        message = "Unknown Error \n.";
    }
    showToaster("error", "Fatal Error", message);
}

function AJAX_Failure(response) {
    showToaster("error", "Fatal Error", "Failure occured");
}