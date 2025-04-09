import { NgModule, ViewChild } from '@angular/core'
import { BrowserModule } from '@angular/platform-browser'
import { Component } from '@angular/core';
import { FileManagerModule, NavigationPaneService, ToolbarService, DetailsViewService, FileManager,FileOpenEventArgs } from '@syncfusion/ej2-angular-filemanager';
import { DocumentEditorContainerModule, ToolbarService as DocToolbarService, DocumentEditorContainerComponent } from '@syncfusion/ej2-angular-documenteditor';
import { PdfViewerModule, PdfViewerComponent } from '@syncfusion/ej2-angular-pdfviewer';
import { SpreadsheetModule, SpreadsheetComponent } from '@syncfusion/ej2-angular-spreadsheet';
import { HttpClientModule, HttpClient } from '@angular/common/http';

@Component({
    imports: [
        HttpClientModule,
        FileManagerModule,
        DocumentEditorContainerModule,
        PdfViewerModule,
        SpreadsheetModule
    ],
    standalone: true,
    selector: 'app-root',
    providers: [NavigationPaneService, ToolbarService, DetailsViewService, DocToolbarService],
    template: `<div class="control-section">
                <ejs-filemanager id='default-filemanager' #fileManager [ajaxSettings]='ajaxSettings' (fileOpen)='onFileOpen($event)'>
                </ejs-filemanager></div>
            <div class="pdf-container hidden container">
                <span class="pdf-title"></span>
                <button (click)="closeView()">Close</button>
                <ejs-pdfviewer id="pdfViewer" #pdfViewer [serviceUrl]="serviceUrl"></ejs-pdfviewer>
            </div>

            <div class="doc-container hidden container">
                <span class="doc-title"></span>
                <button (click)="closeView()">Close</button>
                <ejs-documenteditorcontainer #docEditor [enableToolbar]=true  style="display:block;height:calc(100% - 22px)"></ejs-documenteditorcontainer >
            </div>

            <div class="excel-container hidden container">
                <span class="excel-title"></span>
                <button (click)="closeView()">Close</button>
                <ejs-spreadsheet #spreadsheet [openUrl]="openUrl" ></ejs-spreadsheet>
            </div>`
})
export class AppComponent {
    @ViewChild('fileManager') fileManagerInstance?: FileManager;
    @ViewChild('pdfViewer') pdfViewer!: PdfViewerComponent;
    @ViewChild('docEditor') docEditor!: DocumentEditorContainerComponent;
    @ViewChild('spreadsheet') spreadsheet!: SpreadsheetComponent;
    public hostUrl: string = 'http://localhost:62869/';
    public ajaxSettings: object = {
        url: this.hostUrl + 'api/FileManager/FileOperations',
        getImageUrl: this.hostUrl + 'api/FileManager/GetImage',
        uploadUrl: this.hostUrl + 'api/FileManager/Upload',
        downloadUrl: this.hostUrl + 'api/FileManager/Download',
      };
    serviceUrl='https://ej2services.syncfusion.com/production/web-services/api/pdfviewer'
    fileUrl: string = '';
    docsFilePath: string = '';
    docsFileName: string = '';
    docsFileType: string = '';

    openUrl = 'http://localhost:62869/api/FileManager/OpenExcel';
    constructor(private http: HttpClient) {}

    onFileOpen(event: FileOpenEventArgs| any): void {
        const fileDetails = event.fileDetails;
        const filePath = this.formatFilePath(fileDetails.filterPath, fileDetails.name);

        switch (fileDetails.type) {
            case '.pdf':
                this.showViewer('pdf', fileDetails.name);
                this.getFileStream(filePath, true);
                break;
            case '.docx':
            case '.txt':
            case '.html':
            case '.rtf':
            case '.xml':
                this.docsFilePath = filePath;
                this.showViewer('doc', fileDetails.name);
                this.getFileStream(filePath, false);
                break;
            case '.xlsx':
            case '.csv':
                this.docsFilePath = filePath;
                this.docsFileType = fileDetails.type;
                this.docsFileName = fileDetails.name;
                this.showViewer('excel', fileDetails.name);
                this.getBlob(filePath, fileDetails.name, fileDetails.type);
                break;
            case '.pptx':
                this.showViewer('pdf', fileDetails.name);
                this.convertPptToPdf(filePath, fileDetails.name);
                break;
            default:
                console.warn(`Unsupported file type: ${fileDetails.type}`);
                break;
        }
    
    }
    convertPptToPdf(filePath: string, fileName: string) {
        fetch('http://localhost:62869/api/FileManager/ConvertPptToPdf', {
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
                this.showViewer('pdf', fileName); // Display the PDF
                this.pdfViewer.load(pdfUrl, ''); // Load the PDF in the viewer
            })
            .catch((error) => {
                console.error('Error converting PPT to PDF:', error);
            });
    }
    getFileStream(filePath: string, isPDF: boolean): void {
        const endpoint = `${this.hostUrl}api/FileManager/GetDocument`;
        this.http.post(endpoint, { FileName: filePath, Action: isPDF ? 'LoadPDF' : 'ImportFile' }, { responseType: 'text' })
            .subscribe(response => {
                if (isPDF) {
                    this.pdfViewer.load(response, '');

                } else {
                    this.docEditor.documentEditor.open(response);
                }
            });
    }
    getBlob(filePath: string, fileName: string, fileType: string) {
        let request = new XMLHttpRequest();
        request.responseType = 'blob';
        request.onload = () => {
            let file = new File([request.response], fileName);
            this.spreadsheet.open({ file: file });
        };
        request.open(
            'GET',
            this.hostUrl + 'api/FileManager/GetExcel' + '?FileName=' + filePath
        );
        request.send();
    }

    private formatFilePath(filterPath: string, fileName: string): string {
        return `${filterPath.replace(/\\/g, '/')}${fileName}`;
    }

    private showViewer(type: string, name: string): void {
        document.querySelector('.control-section')?.classList.add('hidden');
        const container = document.querySelector(`.${type}-container`);
        const titleElement = document.querySelector(`.${type}-title`);

        if (container) container.classList.remove('hidden');
        if (titleElement) titleElement.textContent = name;
    }

    closeView(): void {
        document.querySelectorAll('.container').forEach(el => el.classList.add('hidden'));
        document.querySelector('.control-section')?.classList.remove('hidden');
        (this.fileManagerInstance as FileManager).refreshLayout();
    }

}

