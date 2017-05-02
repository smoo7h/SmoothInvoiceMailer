using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SmoothInvoiceMailer
{
    public class GoogleSheetHelper
    {
         
        public GoogleSheetHelper(string applicationName, string jsonSecret, string projectName)
        {
          
            ApplicationName = applicationName;
            ProjectName = projectName;
            JsonFile = jsonSecret;
            service = SetUpSheetsService();
        }

        public string JsonFile;

        public string ProjectName;
        private static string[] Scopes = { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
        private static string ApplicationName;

        public SheetsService service;

        private SheetsService SetUpSheetsService()
        {

            UserCredential credential;

            using (var stream =
                new FileStream(JsonFile, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, String.Format(".credentials/{0}.json",ProjectName));

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        private static int GetNumberFromChar(char cha)
        {
            int index = char.ToUpper(cha) - 65;

            return index;

        }

     
        public void UpdateCellValue(string column, int row, int sheet,string value, string sheetId)
        {
        

            var reqs = new BatchUpdateSpreadsheetRequest();
            reqs.Requests = new List<Request>();
            string[] colNames = new[] { value };

            // Create starting coordinate where data would be written to

            GridCoordinate gc = new GridCoordinate();
            gc.ColumnIndex = GetNumberFromChar(column[0]);
            gc.RowIndex = row-1;
            gc.SheetId = sheet; // Your specific sheet ID here

            Request rq = new Request();
            rq.UpdateCells = new UpdateCellsRequest();
            rq.UpdateCells.Start = gc;
            rq.UpdateCells.Fields = "*"; // needed by API, throws error if null

            // Assigning data to cells
            RowData rd = new RowData();
            List<CellData> lcd = new List<CellData>();
            foreach (String s in colNames)
            {
                ExtendedValue ev = new ExtendedValue();
                ev.StringValue = s;

                CellData cd = new CellData();
                cd.UserEnteredValue = ev;
                lcd.Add(cd);
            }
            rd.Values = lcd;

            // Put cell data into a row
            List<RowData> lrd = new List<RowData>();
            lrd.Add(rd);
            rq.UpdateCells.Rows = lrd;

            // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
            reqs.Requests.Add(rq);

            // Execute request
            BatchUpdateSpreadsheetResponse response = service.Spreadsheets.BatchUpdate(reqs, sheetId).Execute(); // Replace Spreadsheet.SpreadsheetId with your recently created spreadsheet ID



        }
        public void UpdateCellValue(int column,int row,int sheet, string value, string sheetId)
        {

            var reqs = new BatchUpdateSpreadsheetRequest();
            reqs.Requests = new List<Request>();
            string[] colNames = new[] { value };

            // Create starting coordinate where data would be written to

            GridCoordinate gc = new GridCoordinate();
            gc.ColumnIndex = column;
            gc.RowIndex = row ;
            gc.SheetId = sheet; // Your specific sheet ID here

            Request rq = new Request();
            rq.UpdateCells = new UpdateCellsRequest();
            rq.UpdateCells.Start = gc;
            rq.UpdateCells.Fields = "*"; // needed by API, throws error if null

            // Assigning data to cells
            RowData rd = new RowData();
            List<CellData> lcd = new List<CellData>();
            foreach (String s in colNames)
            {
                ExtendedValue ev = new ExtendedValue();
                ev.StringValue = s;

                CellData cd = new CellData();
                cd.UserEnteredValue = ev;
                lcd.Add(cd);
            }
            rd.Values = lcd;

            // Put cell data into a row
            List<RowData> lrd = new List<RowData>();
            lrd.Add(rd);
            rq.UpdateCells.Rows = lrd;

            // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet
            reqs.Requests.Add(rq);

            // Execute request
            BatchUpdateSpreadsheetResponse response = service.Spreadsheets.BatchUpdate(reqs, sheetId).Execute(); // Replace Spreadsheet.SpreadsheetId with your recently created spreadsheet ID

        

        }
    }
}
