using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.Reflection.PortableExecutable;
using System.Xml.Linq;
using iTextSharp.text.pdf;
using System;

namespace PDFFunctions
{
    public class Functions
    {
        const string LIBRE_OFFICE_BIN = "/usr/bin/libreoffice";

        private readonly ILogger<Functions> _logger;
        public Functions(ILogger<Functions> logger)
        {
            _logger = logger;
        }

        [Function("MergePDFDocuments"), RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
        public async Task<IActionResult> RunMerge([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req)
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
                return LogErrorMessage(ex);
            }
        }

        [Function("ConvertDocumentToPdf"), RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
        public async Task<IActionResult> RunConvert([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req)
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
                var file = req.Form.Files[0];
                if (!file.FileName.Contains(".doc", StringComparison.OrdinalIgnoreCase) && !file.FileName.Contains(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    return new BadRequestObjectResult("Incorrect file type. Expected '.docx' or '.xlsx'.");
                }
                using (Stream stream = file.OpenReadStream())
                {
                    SaveStreamAsFile(tempFilePath, stream, file.FileName);
                }
                // Convert file to PDF using libreoffice
                var pdfProcess = new Process();
                pdfProcess.StartInfo.FileName = LIBRE_OFFICE_BIN;
                pdfProcess.StartInfo.Arguments = $"--norestore --nofirststartwizard --headless --convert-to pdf \"{Path.Combine(tempFilePath, file.FileName)}\"";
                pdfProcess.StartInfo.WorkingDirectory = Path.GetDirectoryName(Path.Combine(tempFilePath, file.FileName)); //This is really important
                pdfProcess.Start();
                pdfProcess.WaitForExit();
                //Set the destination file path
                var destinationFileName = $"{Path.Combine(Path.GetDirectoryName(Path.Combine(tempFilePath, file.FileName)), Path.GetFileNameWithoutExtension(Path.Combine(tempFilePath, file.FileName)))}.pdf";
                //Check for exit of process
                if (pdfProcess.ExitCode != 0)
                    throw new Exception("Failed to convert file");
                else
                {
                    int totalChecks = 10;
                    int currentCheck = 1;
                    while (currentCheck <= totalChecks)
                    {
                        if (File.Exists(destinationFileName))
                        {
                            // File conversion was successful
                            break;
                        }
                        Thread.Sleep(500); // LibreOffice doesn't immediately create PDF output once the command is run
                        currentCheck++;
                    }
                }
                // Check if file was converted properly
                if (!File.Exists(destinationFileName))
                {
                    return new BadRequestObjectResult("Error converting file to PDF");
                }
                //Create a new output file stream
                FileStream fStream = new FileStream(destinationFileName, FileMode.Open);
                //Return the file stream
                return new OkObjectResult(fStream);
            }
            catch (Exception ex)
            {
                return LogErrorMessage(ex);
            }
        }
        //Helper method to save file back to file stream
        private void SaveStreamAsFile(string filePath, Stream inputStream, string fileName)
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
        private BadRequestObjectResult LogErrorMessage(Exception ex)
        {
            string errorMessages = $"Index #1\nMessage: {ex.Message}\nStack Trace: {ex.StackTrace}\nSource: {ex.Source}\nTarget Site: {ex.TargetSite}";
            _logger.LogInformation(errorMessages);
            return new BadRequestObjectResult(errorMessages);
        }
    }
}
