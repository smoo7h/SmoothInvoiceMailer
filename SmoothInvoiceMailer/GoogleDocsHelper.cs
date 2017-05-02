using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SmoothInvoiceMailer
{
    public class GoogleDocsHelper
    {
        public DriveService service;
        public GoogleDocsHelper()
        {
            service = SetUpDriveService();
        }
        static string[] Scopes = { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "SmoothInvoiceMailer";

        public string DownloadPDF(string fileId, string fileName)
        {
           
            var request = service.Files.Export(fileId, "application/pdf");
            var stream = new System.IO.MemoryStream();
            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.
            request.MediaDownloader.ProgressChanged +=
                    (IDownloadProgress progress) =>
                    {
                        switch (progress.Status)
                        {
                            case DownloadStatus.Downloading:
                                {
                                    Console.WriteLine(progress.BytesDownloaded);
                                    break;
                                }
                            case DownloadStatus.Completed:
                                {
                                    Console.WriteLine("Download complete.");
                                    break;
                                }
                            case DownloadStatus.Failed:
                                {
                                    Console.WriteLine("Download failed.");
                                    break;
                                }
                        }
                    };
            request.Download(stream);
            string path = Path.GetTempPath();
            Console.WriteLine(path);

            using (var fileStream = File.Create(path+fileName))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fileStream);
            }

            return path + fileName;
        }

        private static DriveService SetUpDriveService()
        {

            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/create.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;


        }

        public void UploadFileToDrive(string file, string folder)
        {
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = file.Split('\\').Last(),
                MimeType = "application/pdf",
                Parents= new List<string>
                {
                    folder
                }
            };
            FilesResource.CreateMediaUpload request;
            using (var stream = new System.IO.FileStream(file,
                                    System.IO.FileMode.Open))
            {
                request = service.Files.Create(
                    fileMetadata, stream, "text/csv");
                request.Fields = "id";
                request.Upload();
            }
            var fi = request.ResponseBody;
            Console.WriteLine("File ID: " + fi.Id);

        }

        public Google.Apis.Drive.v3.Data.File CopyFile( string originFileId, string copyTitle)
        {



            Google.Apis.Drive.v3.Data.File copiedFile = new Google.Apis.Drive.v3.Data.File();
            copiedFile.Name = copyTitle;


            copiedFile = service.Files.Copy(copiedFile, originFileId).Execute();

            // Console.Read();
            Console.WriteLine("New File Created: " + copiedFile.Name);


            return copiedFile;
        }
        public static void ListAllFilesInDrive()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/createdoc.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });




            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {

                    Console.WriteLine("{0} ({1})", file.Name, file.Id);
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }


        }
    }

}
