using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Newtonsoft.Json;

namespace PDFFunction;

public static class PdfFunction
{
    [FunctionName("MergePDFDocuments")]
    [RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
    public static async Task<IActionResult> RunMerge(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        try
        {
            // Initialise Variables
            var tempFilePath = Path.GetTempPath();
            // Clean up Temp Folder before starting
            foreach (var tempfiles in Directory.GetFiles(tempFilePath)) File.Delete(tempfiles);
            //Check if correct content type
            if (!req.ContentType.StartsWith("multipart/form-data", StringComparison.OrdinalIgnoreCase))
                return new BadRequestObjectResult("Incorrect content type. Expected 'multipart/form-data'.");
            //Check if files attached
            if (!req.Form.Files.Any() || req.Form.Files == null || req.Form.Files.Count == 0)
                return new BadRequestObjectResult("No files were uploaded");
            // Get Files
            var files = req.Form.Files;
            //Set output filename
            var outputfilename = $"{tempFilePath}{Guid.NewGuid()}.pdf";
            using (var fstream = new FileStream(outputfilename, FileMode.Create))
            {
                //Create new empty document
                using (var document = new Document(PageSize.A4, 10f, 10f, 10f, 0f))
                {
                    var pdf = new PdfCopy(document, fstream);
                    document.Open();
                    foreach (var file in files)
                    {
                        //Open file in read stream and save as temp file
                        using (var stream = file.OpenReadStream())
                        {
                            SaveStreamAsFile(tempFilePath, stream, file.FileName);
                        }

                        //Open file with reader
                        using (var reader = new PdfReader(Path.Combine(tempFilePath, file.FileName)))
                        {
                            //Add reader contents to pdf
                            pdf.AddDocument(reader);
                        }
                    }
                }
            }

            //Create a new output file stream
            var fStream = new FileStream(outputfilename, FileMode.Open);
            //Return the file stream
            return new OkObjectResult(fStream);
        }
        catch (Exception ex)
        {
            return LogErrorMessage(ex, log);
        }
    }

    [FunctionName("ConvertDocumentToPdf")]
    [RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
    public static async Task<IActionResult> RunConvert(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        // Initialise Variables
        var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
        dynamic jsonInput = JsonConvert.DeserializeObject(requestBody);
        var tempFilePath = Path.GetTempPath();
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
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            var requestUrl =
                $"https://graph.microsoft.com/v1.0/drives/{jsonInput["driveId"]}/items/{jsonInput["fileId"]}/content?format=pdf";
            //Send request
            var response = await httpClient.GetAsync(requestUrl);
            if (response.IsSuccessStatusCode)
            {
                //Read response content and create file
                var pdfContent = await response.Content.ReadAsByteArrayAsync();
                using (var fStream = new FileStream(outputfilename, FileMode.Create))
                {
                    await fStream.WriteAsync(pdfContent, 0, pdfContent.Length);
                    Console.WriteLine("Document converted to PDF successfully.");
                }

                var fileBytes = File.ReadAllBytes(outputfilename);
                return new FileContentResult(fileBytes, "application/pdf")
                {
                    FileDownloadName = "outputfilename.pdf"
                };
            }

            log.LogError(await response.Content.ReadAsStringAsync());
            return new BadRequestObjectResult(await response.Content.ReadAsStringAsync());
        }
        catch (Exception ex)
        {
            return LogErrorMessage(ex, log);
        }
    }

    [FunctionName("AddWatermarkToPdf")]
    [RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
    public static async Task<IActionResult> RunAddWatermark(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        try
        {
            // Initialise Variables
            var tempFilePath = Path.GetTempPath();
            // Clean up Temp Folder before starting
            /*foreach (var tempfiles in Directory.GetFiles(tempFilePath))
            {
                File.Delete(tempfiles);
            }*/
            // Check if correct content type
            if (!req.ContentType.StartsWith("multipart/form-data", StringComparison.OrdinalIgnoreCase))
                return new BadRequestObjectResult("Incorrect content type. Expected 'multipart/form-data'.");
            // Check if files attached
            if (req.Form.Files == null || !req.Form.Files.Any())
                return new BadRequestObjectResult("No files were uploaded.");
            // Check if watermark text is provided
            if (!req.Form.TryGetValue("watermarkText", out var watermarkText) ||
                string.IsNullOrWhiteSpace(watermarkText)) return new BadRequestObjectResult("Missing watermark text.");
            // Get the input PDF file
            var file = req.Form.Files.First();
            var inputFilePath = Path.Combine(tempFilePath, Guid.NewGuid() + "_" + file.FileName);

            SaveStreamAsFile(tempFilePath, file.OpenReadStream(), Path.GetFileName(inputFilePath));

            // Set the output filename
            var outputFilePath = Path.Combine(tempFilePath, $"{file.FileName}_watermarked.pdf");

            // Adding Watermark
            using (var reader = new PdfReader(inputFilePath))
            using (var outputStream = new FileStream(outputFilePath, FileMode.Create))
            {
                using (var stamper = new PdfStamper(reader, outputStream))
                {
                    var pageCount = reader.NumberOfPages;

                    for (var i = 1; i <= pageCount; i++)
                    {
                        var rect = reader.GetPageSize(i);
                        var canvas = stamper.GetOverContent(i);

                        var font = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        canvas.BeginText();
                        canvas.SetFontAndSize(font, 50);
                        canvas.SetColorFill(BaseColor.BLUE);
                        canvas.ShowTextAligned(Element.ALIGN_CENTER, watermarkText, rect.Width / 2, rect.Height / 2,
                            45);
                        canvas.EndText();
                    }
                }
            }

            // Return the watermarked file
            var watermarkedStream = new FileStream(outputFilePath, FileMode.Open);
            return new FileStreamResult(watermarkedStream, "application/pdf")
            {
                FileDownloadName = "watermarked.pdf"
            };
        }
        catch (Exception ex)
        {
            return LogErrorMessage(ex, log);
        }
    }

    //Helper method to save file back to file stream
    private static void SaveStreamAsFile(string filePath, Stream inputStream, string fileName)
    {
        var info = new DirectoryInfo(filePath);
        if (!info.Exists) info.Create();
        var path = Path.Combine(filePath, fileName);
        using (var outputFileStream = new FileStream(path, FileMode.Create))
        {
            inputStream.CopyTo(outputFileStream);
        }
    }

    private static BadRequestObjectResult LogErrorMessage(Exception ex, ILogger log)
    {
        var errorMessages =
            $"Index #1\nMessage: {ex.Message}\nStack Trace: {ex.StackTrace}\nSource: {ex.Source}\nTarget Site: {ex.TargetSite}";
        log.LogError(errorMessages);
        return new BadRequestObjectResult(errorMessages);
    }
}