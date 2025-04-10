import { DetailsView, FileManagerComponent, NavigationPane, Toolbar, Inject } from '@syncfusion/ej2-react-filemanager';
import { PdfViewerComponent } from '@syncfusion/ej2-react-pdfviewer';
import { DocumentEditorContainerComponent } from '@syncfusion/ej2-react-documenteditor';
import { SpreadsheetComponent } from '@syncfusion/ej2-react-spreadsheet';
import * as React from 'react';
import './App.css';

function App() {
    function onfileOpen(args) {

        const fileName = args.fileDetails.name;
        const filePath = args.fileDetails.filterPath.replace(/\\/g, "/") + fileName;
        const fileType = args.fileDetails.type;

        switch (fileType) {
            case ".docx":
            case ".txt":
            case ".html":
            case ".rtf":
            case ".xml":
                showDocViewer( fileName );
                getFileStream(filePath, false);
                break;
            case ".pdf":
                showPDFViewer(fileName);
                getFileStream(filePath, true);
                break;
            case ".csv":
                showExcelViewer( fileName );
                getBlob(filePath, fileName, fileType);
                break;
            case ".pptx":
                showPDFViewer(fileName );
                convertPptToPdf(filePath, fileName);
                break;
            default:
                console.error(`Unsupported file type: ${fileType}`);
        }
    }
    function showDocViewer(name) {
        document.getElementsByClassName("doc-title")[0].innerHTML = name;
        (document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
        (document.getElementsByClassName("doc-container")[0]).style.visibility = "visible";
    }

    //Shows the PDF viewer

    function showPDFViewer(name) {
        (document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
        (document.getElementsByClassName("pdf-container")[0]).style.visibility = "visible";
        document.getElementsByClassName("pdf-title")[0].innerHTML = name;
    }
    function showExcelViewer(name) {
        (document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
        (document.getElementsByClassName("excel-container")[0]).style.visibility = "visible";
        document.getElementsByClassName("excel-title")[0].innerHTML = name;
    }

    // sends HTTPPost Request to read the file as fileStream.

    function getFileStream(filePath, isPDF) {
        let ajax = new XMLHttpRequest();
        ajax.open("POST", hostUrl + "api/FileManager/GetDocument", true);
        ajax.setRequestHeader("content-type", "application/json");
        ajax.onreadystatechange = () => {
            if (ajax.readyState === 4) {
                if (ajax.status === 200 || ajax.status === 304) {
                    if (!isPDF) {
                        docObj.documentEditor.open(ajax.responseText);
                    } else {
                        pdfObj.load(ajax.responseText, null);
                    }
                }
            }
        };
        ajax.send(JSON.stringify({ "FileName": filePath, "Action": (!isPDF ? "ImportFile" : "LoadPDF") }));
    }

    function getBlob(filePath, fileName, fileType) {
        let request = new XMLHttpRequest();
        request.responseType = 'blob';
        request.onload = () => {
            let file = new File([request.response], fileName);
            spreadsheetObj.open({ file: file });
        };
        request.open(
            'GET',
            hostUrl + 'api/FileManager/GetExcel' + '?FileName=' + filePath
        );
        request.send();
    }

    function convertPptToPdf(filePath, fileName) {
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
                const pdfUrl = URL.createObjectURL(blob); // Create a URL for the PDF blob
                showPDFViewer(fileName);
                pdfObj.load(pdfUrl, '');
            })
            .catch((error) => {
                console.error('Error converting PPT to PDF:', error);
            });
    }
    function closeView() {
        (document.getElementsByClassName("file-container")[0]).style.visibility = "visible";
        (document.getElementsByClassName("doc-container")[0]).style.visibility = "hidden";
        (document.getElementsByClassName("pdf-container")[0]).style.visibility = "hidden";
        (document.getElementsByClassName("excel-container")[0]).style.visibility = "hidden";
    }
    let pdfObj;
    let docObj;
    let spreadsheetObj;
    let hostUrl = "http://localhost:{port}/";
    return (<><div className="control-section">
        <div className="file-container">
        <FileManagerComponent id="file"
            ajaxSettings={{
                // Replace the hosted port number in the place of "{port}"
                url: hostUrl + "api/FileManager/FileOperations",
                downloadUrl: hostUrl + 'api/FileManager/Download',
                getImageUrl: hostUrl + "api/FileManager/GetImage",
                uploadUrl: hostUrl + 'api/FileManager/Upload'
            }}
            fileOpen={onfileOpen.bind(this)}>
            <Inject services={[NavigationPane, DetailsView, Toolbar]} />
        </FileManagerComponent>
        </div>
    </div>
    <div className="pdf-container hidden container">
        <span className="pdf-title"></span>
        <button onClick={closeView.bind(this)}>Close</button>
        <PdfViewerComponent id="pdfViewer" ref={s => (pdfObj = s)}
            serviceUrl="https://ej2services.syncfusion.com/production/web-services/api/pdfviewer">
        </PdfViewerComponent>
    </div>
    <div className="doc-container hidden container">
        <span className="doc-title"></span>
        <button onClick={closeView.bind(this)}>Close</button>
        <DocumentEditorContainerComponent id="docEditor" ref={s => (docObj = s)} />
    </div>
    <div className="excel-container hidden container">
        <span className="excel-title"></span>
        <button onClick={closeView.bind(this)}>Close</button>
        <SpreadsheetComponent id="spreadsheet" ref={s => (spreadsheetObj = s)} openUrl="http://localhost:{port}/api/FileManager/OpenExcel"></SpreadsheetComponent>
    </div>
    </>);
}
export default App;
