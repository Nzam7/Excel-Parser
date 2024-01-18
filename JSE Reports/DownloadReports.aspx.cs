using HtmlAgilityPack;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using static JSE_Reports.Repository.JSEReports;

namespace JSE_Reports
{
    public partial class DownloadReports : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        [WebMethod]
        public static string DownloadAllReports()
        {
            try
            {
                // Set up your download folder
                string targetDirectory = @"C:\Users\Nzam Pistis\Downloads\JSEReports"; ;
                FileDownloader downloader = new FileDownloader();

                downloader.DownloadFiles(targetDirectory);

                return "Success";
            }
            catch (Exception ex)
            {
                // Handle the exception
                return "Error: " + ex.Message;
            }
        }

        [WebMethod]
        public static string UploadAllReports()
        {
            try
            {   // Get your download folder
                string targetDirectory = @"C:\Users\Nzam Pistis\Downloads\JSEReports";

                // Processing all Excel files in the target directory
                string[] excelFiles = Directory.GetFiles(targetDirectory, "*.xls");

                FileUploader processor = new FileUploader();
                foreach (string excelFile in excelFiles)
                {
                    processor.ProcessExcelFile(excelFile);
                }

                return "Success";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

        [WebMethod]
        public static async Task<string> UploadAllReportsAsync()
        {
            try
            {
                string targetDirectory = @"C:\Users\Nzam Pistis\Downloads\JSEReports";

                string[] excelFiles = Directory.GetFiles(targetDirectory, "*.xls");

                FileUploaderAsync processor = new FileUploaderAsync();
                foreach (string excelFile in excelFiles)
                {
                    await processor.ProcessExcelFile(excelFile);
                }

                return "Success";
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

    }
}