# PDF Function - Azure Functions

## Overview
This Azure Functions project provides two key functionalities for processing PDF documents:
1. **MergePDFDocuments**: Merges multiple uploaded PDF documents into a single file.
2. **ConvertDocumentToPdf**: Converts documents stored in Microsoft OneDrive or SharePoint to PDF using Microsoft Graph API.

These functions are built using C# and utilize the Azure Functions SDK.

## Features
- **Merge Multiple PDFs**: Allows merging multiple PDF files into a single document.
- **Convert Office Files to PDF**: Uses Microsoft Graph API to convert `.docx`, `.xlsx`, or other supported formats to PDF.
- **Handles Large Files**: Uses `RequestFormLimits` to support large file uploads.
- **Integration with Microsoft Graph API**: Securely authenticates using client credentials for file conversion.
- **Error Logging**: Logs errors using `ILogger`.

## Prerequisites
- **.NET SDK** (Latest version compatible with Azure Functions)
- **Azure Functions Core Tools**
- **Azure Subscription**
- **Microsoft Graph API Access**
  - Registered Azure AD App with required permissions
  - `ClientId`, `TenantId`, and `ClientSecret` for authentication
- **iTextSharp** for PDF processing

## Installation & Setup
### 1. Clone the Repository
```sh
git clone <repository-url>
cd PDFFunction
```

### 2. Install Dependencies
Ensure you have the required NuGet packages installed:
```sh
dotnet add package Microsoft.Azure.WebJobs.Extensions.Http
dotnet add package Microsoft.Identity.Client
dotnet add package iTextSharp
dotnet add package Newtonsoft.Json
```

### 3. Run the Azure Function Locally
```sh
dotnet build
func start
```

## Endpoints
### Merge PDF Documents
**Endpoint:**
```
POST /api/MergePDFDocuments
```
**Request:**
- Content-Type: `multipart/form-data`
- Attach multiple PDF files

**Response:**
- Returns a merged PDF file as a stream

### Convert Document to PDF
**Endpoint:**
```
POST /api/ConvertDocumentToPdf
```
**Request:**
- Content-Type: `application/json`
- JSON Payload:
  ```json
  {
    "ClientId": "<AzureAD_ClientId>",
    "TenantId": "<AzureAD_TenantId>",
    "ClientSecret": "<AzureAD_ClientSecret>",
    "driveId": "<OneDrive_SharePoint_DriveId>",
    "fileId": "<FileId_to_Convert>"
  }
  ```

**Response:**
- Returns a converted PDF file as a response stream

## Error Handling
If an error occurs, the function returns a `400 Bad Request` with detailed error logs captured using `ILogger`.

## Deployment
1. **Publish to Azure**
```sh
dotnet publish -c Release
```
2. **Deploy using Azure CLI**
```sh
az functionapp deployment source config-zip -g <resource-group> -n <function-app-name> --src <zip-file>
```

## Notes
- Temporary files are stored in the system temp directory and deleted after processing.
- The `ConvertDocumentToPdf` function requires Microsoft Graph API permissions and Azure AD authentication.

## License
This project is licensed under the MIT License.

## Author
Developed by [Your Name/Organization]

