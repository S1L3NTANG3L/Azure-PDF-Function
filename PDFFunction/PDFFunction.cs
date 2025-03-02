using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Linq;
using iTextSharp.text.pdf;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Net.Http;
using Newtonsoft.Json;
using System.Reflection.PortableExecutable;

namespace PDFFunction
{
    public static class PDFFunction
    {
        [FunctionName("MergePDFDocuments"), RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
        public static async Task<IActionResult> RunMerge([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
        {
            try
            {
                // Initialise Variables
                string tempFilePath = Path.GetTempPath();
                // Clean up Temp Folder before starting
                foreach (var tempfiles in Directory.GetFiles(tempFilePath))
                {
                    File.Delete(tempfiles);
                }
                //Check if correct content type
                if (!req.ContentType.StartsWith("multipart/form-data", StringComparison.OrdinalIgnoreCase))
                {
                    return new BadRequestObjectResult("Incorrect content type. Expected 'multipart/form-data'.");
                }
                //Check if files attached
                if (!req.Form.Files.Any() || req.Form.Files == null || req.Form.Files.Count == 0)
                {
                    return new BadRequestObjectResult("No files were uploaded");
                }
                // Get Files
                var files = req.Form.Files;
                //Set output filename
                var outputfilename = $"{tempFilePath}{Guid.NewGuid()}.pdf";
                using (FileStream fstream = new FileStream(outputfilename, FileMode.Create))
                {
                    //Create new empty document
                    using (iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 10f, 10f, 10f, 0f))
                    {
                        PdfCopy pdf = new PdfCopy(document, fstream);
                        document.Open();
                        foreach (var file in files)
                        {
                            //Open file in read stream and save as temp file
                            using (Stream stream = file.OpenReadStream())
                            {
                                SaveStreamAsFile(tempFilePath, stream, file.FileName);
                            }
                            //Open file with reader
                            using (PdfReader reader = new PdfReader(Path.Combine(tempFilePath, file.FileName)))
                            {
                                //Add reader contents to pdf
                                pdf.AddDocument(reader);
                            }
                        }
                    }
                }
                //Create a new output file stream
                FileStream fStream = new FileStream(outputfilename, FileMode.Open);
                //Return the file stream
                return new OkObjectResult(fStream);
            }
            catch (Exception ex)
            {
                return LogErrorMessage(ex, log);
            }
        }

        [FunctionName("ConvertDocumentToPdf"), RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
        public static async Task<IActionResult> RunConvert([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
        {
            // Initialise Variables
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic jsonInput = JsonConvert.DeserializeObject(requestBody);
            string tempFilePath = Path.GetTempPath();
            try
            {
                //Set output filename
                var outputfilename = $"{tempFilePath}{Guid.NewGuid()}.pdf";
                //Get json inputs
                string clientId = jsonInput["ClientId"].ToString();
                string tenantId = jsonInput["TenantId"].ToString();
                string clientSecret = jsonInput["ClientSecret"].ToString();
                string[] scopes = { "https://graph.microsoft.com/.default" };
                //Create Confidential Client
                var confidentialClient = ConfidentialClientApplicationBuilder.Create(clientId)
                        .WithClientSecret(clientSecret)
                        .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                        .Build();
                //Get Bearer Token
                var authResult = await confidentialClient.AcquireTokenForClient(scopes).ExecuteAsync();
                //Create HTTP Request
                var httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                var requestUrl = $"https://graph.microsoft.com/v1.0/drives/{jsonInput["driveId"]}/items/{jsonInput["fileId"]}/content?format=pdf";
                //Send request
                var response = await httpClient.GetAsync(requestUrl);
                if (response.IsSuccessStatusCode)
                {
                    //Read response content and create file
                    var pdfContent = await response.Content.ReadAsByteArrayAsync();
                    using (FileStream fStream = new FileStream(outputfilename, FileMode.Create))
                    {
                        await fStream.WriteAsync(pdfContent, 0, pdfContent.Length);
                        Console.WriteLine("Document converted to PDF successfully.");
                    }
                    var fileBytes = System.IO.File.ReadAllBytes(outputfilename);
                    return new FileContentResult(fileBytes, "application/pdf")
                    {
                        FileDownloadName = "outputfilename.pdf"
                    };
                }
                else
                {
                    log.LogError(await response.Content.ReadAsStringAsync());
                    return new BadRequestObjectResult(await response.Content.ReadAsStringAsync());
                }
            }
            catch (Exception ex)
            {
                return LogErrorMessage(ex, log);
            }
        }
        //Helper method to save file back to file stream
        private static void SaveStreamAsFile(string filePath, Stream inputStream, string fileName)
        {
            DirectoryInfo info = new DirectoryInfo(filePath);
            if (!info.Exists)
            {
                info.Create();
            }
            string path = System.IO.Path.Combine(filePath, fileName);
            using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
            {
                inputStream.CopyTo(outputFileStream);
            }
        }
        private static BadRequestObjectResult LogErrorMessage(Exception ex, ILogger log)
        {
            string errorMessages = $"Index #1\nMessage: {ex.Message}\nStack Trace: {ex.StackTrace}\nSource: {ex.Source}\nTarget Site: {ex.TargetSite}";
            log.LogError(errorMessages);
            return new BadRequestObjectResult(errorMessages);
        }
    }
}