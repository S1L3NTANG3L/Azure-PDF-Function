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
            // Clean up Temp Folder before starting (optional, but good for temp files)
            foreach (var tempfiles in Directory.GetFiles(tempFilePath))
            {
                try { File.Delete(tempfiles); } catch (Exception ex) { log.LogWarning($"Could not delete temp file {tempfiles}: {ex.Message}"); }
            }

            // Check if correct content type
            if (!req.ContentType.StartsWith("multipart/form-data", StringComparison.OrdinalIgnoreCase))
                return new BadRequestObjectResult("Incorrect content type. Expected 'multipart/form-data'.");
            // Check if files attached
            if (req.Form.Files == null || !req.Form.Files.Any())
                return new BadRequestObjectResult("No files were uploaded.");

            // Get watermark text
            if (!req.Form.TryGetValue("watermarkText", out var watermarkText) ||
                string.IsNullOrWhiteSpace(watermarkText))
                return new BadRequestObjectResult("Missing watermark text.");

            
            req.Form.TryGetValue("watermarkColor", out var watermarkColorString);
           
            req.Form.TryGetValue("watermarkOpacity", out var watermarkOpacityString);

            BaseColor watermarkColor = ParseColor(watermarkColorString, watermarkOpacityString, log); // Default to Red with default opacity

            req.Form.TryGetValue("watermarkFont", out var watermarkFontString);
            BaseFont watermarkFont = ParseFont(watermarkFontString, log); // Default to Helvetica if not provided or invalid

            // Get the input PDF file
            var file = req.Form.Files.First();
            var inputFilePath = Path.Combine(tempFilePath, Guid.NewGuid() + "_" + file.FileName);

            // Save the uploaded file to a temporary location
            using (var stream = file.OpenReadStream())
            {
                SaveStreamAsFile(tempFilePath, stream, Path.GetFileName(inputFilePath));
            }

            // Set the output filename
            var outputFilePath = Path.Combine(tempFilePath, $"{Guid.NewGuid()}_{file.FileName}_watermarked.pdf");

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

                        canvas.BeginText();
                        // Adjust font size for better fit as seen in the screenshot
                        canvas.SetFontAndSize(watermarkFont,35); // Increased font size
                        canvas.SetColorFill(watermarkColor); // Use the parsed color with opacity

                        // Calculate text width to center it
                        float textWidth = watermarkFont.GetWidthPoint(watermarkText, 45); // Use the same font size

                    
                        float xPos = rect.Width * 0.50f; // Roughly 1/4th from left
                        float yPos = rect.Height * 0.60f; // Roughly 1/4th from bottom

                        canvas.ShowTextAligned(Element.ALIGN_CENTER, watermarkText, xPos, yPos, 30); // Adjusted position
                        canvas.EndText();
                    }
                }
            }

            // Read the watermarked file into a byte array to return
            var watermarkedFileBytes = await File.ReadAllBytesAsync(outputFilePath);

            // Clean up temporary files after processing
            try
            {
                File.Delete(inputFilePath);
                File.Delete(outputFilePath);
            }
            catch (Exception ex)
            {
                log.LogWarning($"Error deleting temporary files: {ex.Message}");
            }

            // Return the watermarked file
            return new FileContentResult(watermarkedFileBytes, "application/pdf")
            {
                FileDownloadName = $"watermarked_{file.FileName}"
            };
        }
        catch (Exception ex)
        {
            return LogErrorMessage(ex, log);
        }
    }

 
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


    private static BaseColor ParseColor(string colorString, string opacityString, ILogger log)
    {
        // Default opacity if not provided or invalid
        float opacity = 0.3f; // Default to 30% opacity
        if (!string.IsNullOrWhiteSpace(opacityString) && float.TryParse(opacityString, out float parsedOpacity))
        {
            opacity = Math.Max(0.0f, Math.Min(1.0f, parsedOpacity)); // Clamp between 0 and 1
        }
        else if (!string.IsNullOrWhiteSpace(opacityString))
        {
            log.LogWarning($"Could not parse opacity string '{opacityString}', defaulting to {opacity * 100}%.");
        }
        else
        {
            log.LogInformation($"No watermark opacity provided, defaulting to {opacity * 100}%.");
        }

        BaseColor baseColor = BaseColor.RED; // Default color

        if (string.IsNullOrWhiteSpace(colorString))
        {
            log.LogInformation("No watermark color provided, defaulting to Red.");
        }
        else
        {
            // Try parsing as a known color name
            switch (colorString.ToLowerInvariant())
            {
                case "black": baseColor = BaseColor.BLACK; break;
                case "blue": baseColor = BaseColor.BLUE; break;
                case "cyan": baseColor = BaseColor.CYAN; break;
                case "darkgray": baseColor = BaseColor.DARK_GRAY; break;
                case "gray": baseColor = BaseColor.GRAY; break;
                case "green": baseColor = BaseColor.GREEN; break;
                case "lightgray": baseColor = BaseColor.LIGHT_GRAY; break;
                case "magenta": baseColor = BaseColor.MAGENTA; break;
                case "orange": baseColor = BaseColor.ORANGE; break;
                case "pink": baseColor = BaseColor.PINK; break;
                case "red": baseColor = BaseColor.RED; break;
                case "white": baseColor = BaseColor.WHITE; break;
                case "yellow": baseColor = BaseColor.YELLOW; break;
                default:
                    // Try parsing as a hex color code (e.g., #RRGGBB or RRGGBB)
                    try
                    {
                        string hexColor = colorString.StartsWith("#") ? colorString.Substring(1) : colorString;
                        if (hexColor.Length == 6)
                        {
                            int r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                            int g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                            int b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                            baseColor = new BaseColor(r, g, b);
                        }
                        else
                        {
                            log.LogWarning($"Unsupported watermark color format '{colorString}', defaulting to Red.");
                        }
                    }
                    catch (Exception ex)
                    {
                        log.LogWarning($"Could not parse color string '{colorString}'. Error: {ex.Message}. Defaulting to Red.");
                    }
                    break;
            }
        }

        // Apply opacity to the chosen base color
        return new BaseColor(baseColor.R, baseColor.G, baseColor.B, (int)(opacity * 255));
    }

    private static BaseFont ParseFont(string fontString, ILogger log)
    {
        if (string.IsNullOrWhiteSpace(fontString))
        {
            log.LogInformation("No watermark font provided, defaulting to Helvetica.");
            return BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        }

        try
        {
            switch (fontString.ToLowerInvariant())
            {
                case "helvetica":
                    return BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "helvetica-bold":
                    return BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "helvetica-oblique":
                    return BaseFont.CreateFont(BaseFont.HELVETICA_OBLIQUE, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "helvetica-boldoblique":
                    return BaseFont.CreateFont(BaseFont.HELVETICA_BOLDOBLIQUE, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "times-roman":
                    return BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "times-bold":
                    return BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "times-italic":
                    return BaseFont.CreateFont(BaseFont.TIMES_ITALIC, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "times-bolditalic":
                    return BaseFont.CreateFont(BaseFont.TIMES_BOLDITALIC, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "courier":
                    return BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "courier-bold":
                    return BaseFont.CreateFont(BaseFont.COURIER_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "courier-oblique":
                    return BaseFont.CreateFont(BaseFont.COURIER_OBLIQUE, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "courier-boldoblique":
                    return BaseFont.CreateFont(BaseFont.COURIER_BOLDOBLIQUE, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "symbol":
                    return BaseFont.CreateFont(BaseFont.SYMBOL, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                case "zapfdingbats":
                    return BaseFont.CreateFont(BaseFont.ZAPFDINGBATS, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                default:
                    log.LogWarning($"Unsupported watermark font '{fontString}', defaulting to Helvetica.");
                    return BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            }
        }
        catch (DocumentException ex)
        {
            log.LogError($"Error creating font '{fontString}': {ex.Message}. Defaulting to Helvetica.");
            return BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
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
