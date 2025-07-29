# Azure PDF Function

A snazzy set of Azure Functions for automating your PDF-tinkering urges: merging, converting, and watermarking PDF documents using .NET, iTextSharp, and Microsoft Graph. Because we all know manual PDF editing is so last decade.

---

## Features

- **MergePDFDocuments**: Combine multiple PDF files into a single mighty PDF.
- **ConvertDocumentToPdf**: Convert documents from Microsoft OneDrive/SharePoint (via Microsoft Graph) into that universally misunderstood format: PDF.
- **AddWatermarkToPdf**: Slap a custom watermark onto your PDFs and assert dominance.

## Prerequisites

- .NET Core & Azure Functions runtime
- Microsoft Graph API access (for conversion)
- iTextSharp library
- Microsoft.Identity.Client, Newtonsoft.Json, and all the usual suspects

## Usage

### Merging PDFs

**POST** to `/api/MergePDFDocuments`  
Content-Type: `multipart/form-data`  
Attach your PDF files.

### Converting Office Files to PDF

**POST** to `/api/ConvertDocumentToPdf`  
Content-Type: `application/json`  
Body example:
```json
{
  "ClientId": "<azure-app-client-id>",
  "TenantId": "<azure-tenant-id>",
  "ClientSecret": "<azure-app-client-secret>",
  "driveId": "<drive-id>",
  "fileId": "<file-id>"
}
```

### Watermarking PDFs

**POST** to `/api/AddWatermarkToPdf`  
Content-Type: `multipart/form-data`  
Fields:
- `watermarkText`: The text to haunt your PDF.
- `watermarkColor`: (optional) Named colour or hex, e.g. `#FF0000`.
- `watermarkOpacity`: (optional) A float between 0 and 1.
- `watermarkFont`: (optional) Eg: `helvetica`, `times-bold`, etc.
- File: The PDF to bless with your watermark.

## Local Development

1. Clone repository.
2. Restore NuGet packages.
3. Set up your Azure credentials for the convert function.
4. Run locally with VS Code, Visual Studio, or Azure Functions Core Tools.
5. POST with your favourite REST client—no one’s judging.

## Error Handling

Because things go wrong, errors are logged and returned with useful messages. Read them. Or ignore at your peril.

## License

MIT, because open source should be free—like all the PDFs you're manipulating now.

---

*For more details, consult the code. It’s quite chatty.*
