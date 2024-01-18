<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="DownloadReports.aspx.cs" Inherits="JSE_Reports.DownloadReports" %>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link rel="icon" href="Content/Images/favi.png" type="image/png">
    <title>JSE Report Manager</title>
    <!-- Bootstrap CSS -->
    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <link href="Content/toastr.css" rel="stylesheet" />

    <!-- Custom CSS -->
    <link href="Content/App/Dashboard.css" rel="stylesheet" />
</head>

<body>
    <div class="d-flex justify-content-center align-items-center" style="min-height: 100vh;">
        <button id="DownloadReports" type="button" class="btn btn-primary btn-lg m-2">Download JSE Reports</button>
        <button style="display: none" id="UploadReports" type="button" class="btn btn-primary btn-lg m-2">Upload JSE Reports</button>
    </div>


    <!-- JavaScript -->
    <script src="Scripts/jquery-3.7.0.min.js"></script>
    <script src="Scripts/bootstrap.bundle.js"></script>
    <script src="Scripts/toastr.js"></script>
    <script src="Scripts/App/ToastrExec.js"></script>
    <script src="Scripts/App/AjaxHelper.js"></script>
    <script src="Scripts/App/DownloadReports.js"></script>
</body>
</html>
