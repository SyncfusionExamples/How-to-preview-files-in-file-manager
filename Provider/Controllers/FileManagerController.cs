using Syncfusion.EJ2.FileManager.PhysicalFileProvider;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using Syncfusion.EJ2.FileManager.Base;
using Newtonsoft.Json;
using System.IO;
using System.Text.Json;
using Syncfusion.EJ2.Spreadsheet;
using Syncfusion.XlsIO;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using FormatType = Syncfusion.EJ2.DocumentEditor.FormatType;

namespace EJ2APIServices.Controllers
{

    [Route("api/[controller]")]
    [EnableCors("AllowAllOrigins")]
    public class FileManagerController : Controller
    {
        public PhysicalFileProvider operation;
        public string basePath;
        string root = "wwwroot\\Files";
        public FileManagerController(IWebHostEnvironment hostingEnvironment)
        {
            this.basePath = hostingEnvironment.ContentRootPath;
            this.operation = new PhysicalFileProvider();
            this.operation.RootFolder(this.basePath + "\\" + this.root);
        }
        [Route("FileOperations")]
        public object FileOperations([FromBody] FileManagerDirectoryContent args)
        {
            if (args.Action == "delete" || args.Action == "rename")
            {
                if ((args.TargetPath == null) && (args.Path == ""))
                {
                    FileManagerResponse response = new FileManagerResponse();
                    response.Error = new ErrorDetails { Code = "401", Message = "Restricted to modify the root folder." };
                    return this.operation.ToCamelCase(response);
                }
            }
            switch (args.Action)
            {
                case "read":
                    // reads the file(s) or folder(s) from the given path.
                    return this.operation.ToCamelCase(this.operation.GetFiles(args.Path, args.ShowHiddenItems));
                case "delete":
                    // deletes the selected file(s) or folder(s) from the given path.
                    this.operation.Response = Response;
                    return this.operation.ToCamelCase(this.operation.Delete(args.Path, args.Names));
                case "copy":
                    // copies the selected file(s) or folder(s) from a path and then pastes them into a given target path.
                    return this.operation.ToCamelCase(this.operation.Copy(args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData));
                case "move":
                    // cuts the selected file(s) or folder(s) from a path and then pastes them into a given target path.
                    return this.operation.ToCamelCase(this.operation.Move(args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData));
                case "details":
                    // gets the details of the selected file(s) or folder(s).
                    return this.operation.ToCamelCase(this.operation.Details(args.Path, args.Names, args.Data));
                case "create":
                    // creates a new folder in a given path.
                    return this.operation.ToCamelCase(this.operation.Create(args.Path, args.Name));
                case "search":
                    // gets the list of file(s) or folder(s) from a given path based on the searched key string.
                    return this.operation.ToCamelCase(this.operation.Search(args.Path, args.SearchString, args.ShowHiddenItems, args.CaseSensitive));
                case "rename":
                    // renames a file or folder.
                    return this.operation.ToCamelCase(this.operation.Rename(args.Path, args.Name, args.NewName, false, args.ShowFileExtension, args.Data));
            }
            return null;
        }

        // uploads the file(s) into a specified path
        [Route("Upload")]
        [DisableRequestSizeLimit]
        public IActionResult Upload(string path, long size, IList<IFormFile> uploadFiles, string action)
        {
            try
            {
                FileManagerResponse uploadResponse;
                foreach (var file in uploadFiles)
                {
                    var folders = (file.FileName).Split('/');
                    // checking the folder upload
                    if (folders.Length > 1)
                    {
                        for (var i = 0; i < folders.Length - 1; i++)
                        {
                            string newDirectoryPath = Path.Combine(this.basePath + path, folders[i]);
                            if (Path.GetFullPath(newDirectoryPath) != (Path.GetDirectoryName(newDirectoryPath) + Path.DirectorySeparatorChar + folders[i]))
                            {
                                throw new UnauthorizedAccessException("Access denied for Directory-traversal");
                            }
                            if (!Directory.Exists(newDirectoryPath))
                            {
                                this.operation.ToCamelCase(this.operation.Create(path, folders[i]));
                            }
                            path += folders[i] + "/";
                        }
                    }
                }
                uploadResponse = operation.Upload(path, uploadFiles, action, size, null);
                if (uploadResponse.Error != null)
                {
                    Response.Clear();
                    Response.ContentType = "application/json; charset=utf-8";
                    Response.StatusCode = Convert.ToInt32(uploadResponse.Error.Code);
                    Response.HttpContext.Features.Get<IHttpResponseFeature>().ReasonPhrase = uploadResponse.Error.Message;
                }
            }
            catch (Exception e)
            {
                ErrorDetails er = new ErrorDetails();
                er.Message = e.Message.ToString();
                er.Code = "417";
                er.Message = "Access denied for Directory-traversal";
                Response.Clear();
                Response.ContentType = "application/json; charset=utf-8";
                Response.StatusCode = Convert.ToInt32(er.Code);
                Response.HttpContext.Features.Get<IHttpResponseFeature>().ReasonPhrase = er.Message;
                return Content("");
            }
            return Content("");
        }

        // downloads the selected file(s) and folder(s)
        [Route("Download")]
        public IActionResult Download(string downloadInput)
        {
            FileManagerDirectoryContent args = JsonConvert.DeserializeObject<FileManagerDirectoryContent>(downloadInput);
            return operation.Download(args.Path, args.Names, args.Data);
        }

        // gets the image(s) from the given path
        [Route("GetImage")]
        public IActionResult GetImage(FileManagerDirectoryContent args)
        {
            return this.operation.GetImage(args.Path, args.Id,false,null, null);
        }

        [HttpPost]
        [Route("ConvertPptToPdf")]
        public IActionResult ConvertPptToPdf([FromBody] FileRequest fileRequest)
        {
            if (fileRequest == null || string.IsNullOrEmpty(fileRequest.FilePath))
            {
                return BadRequest("Invalid file path. Please provide a valid file path.");
            }
            string filePath = this.basePath + "\\wwwroot\\Files" + (fileRequest.FilePath).Replace("/", "\\");

            if (!System.IO.File.Exists(filePath) || Path.GetExtension(filePath).ToLower() != ".pptx")
            {
                return BadRequest("Invalid file. Please ensure the file exists and is a PPTX file.");
            }

            try
            {
                using (FileStream pptStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Load the presentation from the file stream
                    using (IPresentation presentation = Presentation.Open(pptStream))
                    {
                        // Convert the presentation to PDF
                        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation))
                        {
                            // Save the PDF to a memory stream
                            using (MemoryStream pdfStream = new MemoryStream())
                            {
                                pdfDocument.Save(pdfStream);
                                pdfStream.Position = 0;

                                // Return the PDF file as a response
                                return File(pdfStream.ToArray(), "application/pdf", "ConvertedPresentation.pdf");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error converting PPT to PDF: {ex.Message}");
            }
        }

        // Create a model to capture the file path from the request
        public class FileRequest
        {
            public string FilePath { get; set; }
        }
        [Route("GetDocument")]
        public string GetDocument([FromBody] CustomParams param)
        {
            string path = this.basePath + "\\wwwroot\\Files" + (param.FileName).Replace("/", "\\");
            if (param.Action == "LoadPDF")
            {
                //for PDF Files
                var docBytes = System.IO.File.ReadAllBytes(path);
                //we can convert the document stream to bytes then convert to base64
                string docBase64 = "data:application/pdf;base64," + Convert.ToBase64String(docBytes);
                return (docBase64);
            }
            else
            {
                //for Doc Files
                try
                {
                    Stream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.ReadWrite);
                    int index = param.FileName.LastIndexOf('.');
                    string type = index > -1 && index < param.FileName.Length - 1 ?
                        param.FileName.Substring(index) : ".docx";
                    WordDocument document = WordDocument.Load(stream, GetFormatType(type.ToLower()));
                    string json = JsonConvert.SerializeObject(document);
                    document.Dispose();
                    stream.Dispose();
                    return json;
                }
                catch
                {
                    return "Failure";
                }
            }
        }
        [Route("SaveDocument")]
        [HttpPost]
        public IActionResult SaveDocument([FromForm] IFormFile data, [FromForm] string fileName)
        {
            if (data == null || string.IsNullOrEmpty(fileName))
            {
                return BadRequest("Invalid file data or filename.");
            }

            try
            {
                string path = this.basePath + "\\wwwroot\\Files" + (fileName).Replace("/", "\\");

                string directory = Path.GetDirectoryName(path);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    data.CopyTo(fileStream);
                }

                return Ok("File saved successfully.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error saving file: {ex.Message}");
            }
        }
        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
            {
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }

            switch (format.ToLower())
            {
                case ".dotx":
                case ".docx":
                case ".docm":
                case ".dotm":
                    return FormatType.Docx;
                case ".dot":
                case ".doc":
                    return FormatType.Doc;
                case ".rtf":
                    return FormatType.Rtf;
                case ".txt":
                    return FormatType.Txt;
                case ".xml":
                    return FormatType.WordML;
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        //For Excel Files
        [Route("GetExcel")]
        public IActionResult GetExcel(CustomParams param)
        {
            string fullPath = this.basePath + "\\wwwroot\\Files" + (param.FileName).Replace("/", "\\");
            FileStream fileStreamInput = new FileStream(fullPath, FileMode.Open, FileAccess.Read);
            FileStreamResult fileStreamResult = new FileStreamResult(fileStreamInput, "APPLICATION/octet-stream");
            return fileStreamResult;
        }
        [Route("OpenExcel")]
        public IActionResult Open(IFormCollection openRequest)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            Stream memStream = (openRequest.Files[0] as IFormFile).OpenReadStream();
            IFormFile formFile = new FormFile(memStream, 0, memStream.Length, "", openRequest.Files[0].FileName); // converting MemoryStream to IFormFile
            OpenRequest open = new OpenRequest();
            open.File = formFile;
            var result = Workbook.Open(open);
            memStream.Close();
            return Content(result);
        }
    }
    public class CustomParams
    {
        public string FileName { get; set; }
        public string Action { get; set; }
    }

}
