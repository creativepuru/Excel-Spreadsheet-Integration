using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel_Spreadsheet_Integration
{
    public partial class Form1 : Form
    {
        private SheetsService sheetsService;
        public Form1()
        {
            InitializeComponent();
            sheetsService = GoogleSheetsAuthentication.GetService();
        }

        private void buttonRead_Click(object sender, EventArgs e)
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            string localExcelFilePath = Path.Combine(currentDirectory, "Spreadsheet_Integration.xlsx");

            List<List<object>> excelData = ReadLocalExcelFile(localExcelFilePath);

            if (excelData.Count > 0)
            {
                string message = "Excel Data:\n\n";
                foreach (var rowData in excelData)
                {
                    message += string.Join("\t", rowData) + "\n";
                }

                MessageBox.Show(message, "Local Excel Data");
            }
            else
            {
                MessageBox.Show("No data found in the local Excel file.", "Local Excel Data");
            }
        }

        private void buttonWrite_Click(object sender, EventArgs e)
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            string localExcelFilePath = Path.Combine(currentDirectory, "Spreadsheet_Integration.xlsx");

            string spreadsheetId = "1YPpeervlTXZgNIrMDbYFpCBPtnkZgxYIpKIKO4Dix2s";
            string range = "Sheet1!A1:Z1000";

            List<List<object>> excelData = ReadLocalExcelFile(localExcelFilePath);
            WriteToGoogleSpreadsheet(spreadsheetId, range, excelData);
        }

        private List<List<object>> ReadLocalExcelFile(string excelFilePath)
        {
            List<List<object>> excelData = new List<List<object>>();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];

            int rowCount = excelWorksheet.UsedRange.Rows.Count;
            int colCount = excelWorksheet.UsedRange.Columns.Count;

            for (int row = 1; row <= rowCount; row++)
            {
                List<object> rowData = new List<object>();
                for (int col = 1; col <= colCount; col++)
                {
                    Excel.Range cell = excelWorksheet.Cells[row, col];
                    rowData.Add(cell.Value);
                }
                excelData.Add(rowData);
            }

            // Close and release Excel objects
            excelWorkbook.Close(false);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(excelApp);

            return excelData;
        }

        private void WriteToGoogleSpreadsheet(string spreadsheetId, string range, List<List<object>> data)
        {
            var valueRange = new ValueRange
            {
                Values = data.Select(row => row.Select(cell => cell).ToList()).ToList<IList<object>>()
            };

            var updateRequest = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            var updateResponse = updateRequest.Execute();
            MessageBox.Show("Data written to Google Spreadsheet successfully!");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }

    public class GoogleSheetsAuthentication
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "naam";
        static string ClientSecretFilePath;

        static GoogleSheetsAuthentication()
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            ClientSecretFilePath = Path.Combine(currentDirectory, "client_secret_290780112564-3fehqenc4flt3j40n34v84t5bveq9l71.apps.googleusercontent.com.json");
        }

        public static SheetsService GetService()
        {
            UserCredential credential;

            using (var stream = new FileStream(ClientSecretFilePath, FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets, Scopes, "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Creating Google Sheets service using the authorized credential
            return new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
    }
}
