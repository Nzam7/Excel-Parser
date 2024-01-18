//Types:
/////success
/////info
/////warning
/////error

function showToaster(type, header, message, timeOut) {
    timeOut = timeOut === undefined ? "3000" : timeOut;
    toastr.clear();
    toastr.options = {
        "closeButton": false,
        "debug": false,
        "newestOnTop": true,
        "progressBar": true,
        "positionClass": "toast-top-right",
        "preventDuplicates": false,
        "onclick": null,
        "showDuration": "200",
        "hideDuration": "200",
        "timeOut": timeOut,
        "extendedTimeOut": "300",
        "showEasing": "swing",
        "hideEasing": "swing",
        "showMethod": "show",
        "hideMethod": "hide"
    };
    toastr[type](message, header);
}


function showToasterButton(type, header, message) {
    toastr.clear();
    toastr.options = {
        "closeButton": false,
        "debug": false,
        "newestOnTop": true,
        "progressBar": true,
        "positionClass": "toast-top-right",
        "preventDuplicates": false,
        "onclick": null,
        "showDuration": "500",
        "hideDuration": "500",
        "timeOut": "0",
        "extendedTimeOut": "1000",
        "showEasing": "swing",
        "hideEasing": "swing",
        "showMethod": "show",
        "hideMethod": "hide"
    };
    toastr[type](message, header);
}