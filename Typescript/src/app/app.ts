import { FileManager, Toolbar, NavigationPane, DetailsView, FileOpenEventArgs } from '@syncfusion/ej2-filemanager';
import { PdfViewer } from '@syncfusion/ej2-pdfviewer'
import { DocumentEditorContainer } from '@syncfusion/ej2-documenteditor';
import { Spreadsheet } from '@syncfusion/ej2-spreadsheet';

FileManager.Inject(Toolbar, NavigationPane, DetailsView);

const hostUrl: string = 'http://localhost:{port}/';

const filemanagerInstance: FileManager = new FileManager({
    ajaxSettings: {
        url: hostUrl + "api/FileManager/FileOperations",
        downloadUrl: hostUrl + "api/FileManager/Download",
        uploadUrl: hostUrl + "api/FileManager/Upload",
        getImageUrl: hostUrl + "api/FileManager/GetImage"
    },
    fileOpen: (args: FileOpenEventArgs | any) => {
        let fileName: string = args.fileDetails.name;
        let filePath: string = args.fileDetails.filterPath.replace(/\\/g, "/") + fileName;
        let fileType: string = args.fileDetails.type;
        switch (fileType) {
            case ".docx":
            case ".txt":
            case ".html":
            case ".rtf":
            case ".xml":
                showDocViewer(fileName);
                getFileStream(filePath, false);
                break;
            
            case ".pdf":
                showPDFViewer(fileName);
                getFileStream(filePath, true);
                break;
            
            case ".csv":
                showExcelViewer(fileName);
                getBlob(filePath, fileName, fileType);
                break;
            
            case ".pptx":
                showPDFViewer(fileName);
                convertPptToPdf(filePath, fileName);
                break;
            
            default:
                console.log("Unsupported file type");
                break;
        }
    }
});

filemanagerInstance.appendTo('#filemanager');

let viewer: PdfViewer = new PdfViewer();

viewer.serviceUrl = 'https://ej2services.syncfusion.com/production/web-services/api/pdfviewer';

viewer.appendTo('#pdfViewer');

let container: DocumentEditorContainer = new DocumentEditorContainer({ height: '590px' });

container.appendTo('#container');

let spreadsheet: Spreadsheet = new Spreadsheet();
spreadsheet.openUrl = 'http://localhost:{port}/api/FileManager/OpenExcel';
spreadsheet.appendTo('#excelViewer');

//Shows the Document viewer

function showDocViewer(name: string) {
    document.getElementsByClassName("doc-title")[0].innerHTML = name;
    (<HTMLElement>document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
    (<HTMLElement>document.getElementsByClassName("doc-container")[0]).style.visibility = "visible";
}

//Shows the PDF viewer

function showPDFViewer(name: string) {
    (<HTMLElement>document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
    (<HTMLElement>document.getElementsByClassName("pdf-container")[0]).style.visibility = "visible";
    document.getElementsByClassName("pdf-title")[0].innerHTML = name;
}
function showExcelViewer(name: string) {
    (<HTMLElement>document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
    (<HTMLElement>document.getElementsByClassName("excel-container")[0]).style.visibility = "visible";
    document.getElementsByClassName("excel-title")[0].innerHTML = name;
}

// sends HTTPPost Request to read the file as fileStream.

function getFileStream(filePath: string, isPDF: boolean) {
    let ajax: XMLHttpRequest = new XMLHttpRequest();
    ajax.open("POST", hostUrl + "api/FileManager/GetDocument", true);
    ajax.setRequestHeader("content-type", "application/json");
    ajax.onreadystatechange = () => {
        if (ajax.readyState === 4) {
            if (ajax.status === 200 || ajax.status === 304) {
                if (!isPDF) {
                    container.documentEditor.open(ajax.responseText);
                } else {
                    var pdfviewer = (<any>document.getElementById('pdfViewer')).ej2_instances[0];
                    pdfviewer.load(ajax.responseText, null);
                }
            }
        }
    };
    ajax.send(JSON.stringify({ "FileName": filePath, "Action": (!isPDF ? "ImportFile" : "LoadPDF") }));
}

function getBlob(filePath: string, fileName: string, fileType: string) {
    let request = new XMLHttpRequest();
    request.responseType = 'blob';
    request.onload = () => {
        let file = new File([request.response], fileName);
        spreadsheet.open({ file: file });
    };
    request.open(
        'GET',
        hostUrl + 'api/FileManager/GetExcel' + '?FileName=' + filePath
    );
    request.send();
}

function convertPptToPdf(filePath: string, fileName: string) {
    fetch('http://localhost:{port}/api/FileManager/ConvertPptToPdf', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ FilePath: filePath }), // Send file path to backend
    })
        .then((response) => {
            if (!response.ok) {
                throw new Error('Failed to convert PPT to PDF');
            }
            return response.blob(); // Get the PDF blob
        })
        .then((blob) => {
            var pdfviewer = (<any>document.getElementById('pdfViewer')).ej2_instances[0];
            const pdfUrl = URL.createObjectURL(blob); // Create a URL for the PDF blob
            showPDFViewer(fileName);
            pdfviewer.load(pdfUrl, '');
        })
        .catch((error) => {
            console.error('Error converting PPT to PDF:', error);
        });
}
