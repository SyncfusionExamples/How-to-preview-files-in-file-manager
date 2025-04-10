<template>
<div id="app">
    <div class="control-section">
    <div class="file-container">
     <ejs-filemanager id="overview_filemanager" :ajaxSettings="ajaxSettings" :fileOpen='onfileOpen'>
     </ejs-filemanager>
    </div>
    <div class="pdf-container hidden container">
      <span class="pdf-title"></span>
      <button v-on:click="closeView">Close</button>
      <ejs-pdfviewer id="pdfViewer" ref="pdfViewer" :serviceUrl="serviceUrl"></ejs-pdfviewer>
    </div>

    <div class="doc-container hidden container">
      <span class="doc-title"></span>
      <button v-on:click="closeView" >Close</button>
      <ejs-documenteditorcontainer id="docViewer" ref="docViewer"></ejs-documenteditorcontainer >
    </div>

    <div class="excel-container hidden container">
      <span class="excel-title"></span>
      <button v-on:click="closeView" >Close</button>
      <ejs-spreadsheet id="excelViewer"  ref="excelViewer" :openUrl="openUrl"></ejs-spreadsheet>
    </div>
</div>
</div>
</template>
<script>
import { FileManagerComponent, DetailsView, NavigationPane, Toolbar } from "@syncfusion/ej2-vue-filemanager";
import { PdfViewerComponent } from "@syncfusion/ej2-vue-pdfviewer";
import { DocumentEditorContainerComponent } from "@syncfusion/ej2-vue-documenteditor";
import { SpreadsheetComponent } from "@syncfusion/ej2-vue-spreadsheet";
import { enableRipple } from "@syncfusion/ej2-base";

enableRipple(true);
let hostUrl = 'http://localhost:{port}/';

export default {
    name: "App",
    components: {
        "ejs-filemanager":FileManagerComponent,
        "ejs-pdfviewer":PdfViewerComponent,
        "ejs-documenteditorcontainer":DocumentEditorContainerComponent,
        "ejs-spreadsheet":SpreadsheetComponent
    },
     data () {
      return {
        ajaxSettings:
        {
          url: hostUrl + 'api/FileManager/FileOperations',
          getImageUrl: hostUrl + 'api/FileManager/GetImage',
          uploadUrl: hostUrl + 'api/FileManager/Upload',
          downloadUrl: hostUrl + 'api/FileManager/Download'
        },
        serviceUrl: 'https://ej2services.syncfusion.com/production/web-services/api/pdfviewer',
        openUrl : 'http://localhost:{port}/api/FileManager/OpenExcel'
      }
    },
    methods: {
      onfileOpen: function (args) {
        let fileName = args.fileDetails.name;
        let filePath = args.fileDetails.filterPath.replace(/\\/g, "/") + fileName;
        let fileType = args.fileDetails.type;
        switch (fileType) {
            case ".docx":
            case ".txt":
            case ".html":
            case ".rtf":
            case ".xml":
                this.showDocViewer(fileName);
                this.getFileStream(filePath, false);
                break;

            case ".pdf":
                this.showPDFViewer(fileName);
                this.getFileStream(filePath, true);
                break;

            case ".csv":
                this.showExcelViewer(fileName);
                this.getBlob(filePath, fileName, fileType);
                break;

            case ".pptx":
                this.showPDFViewer(fileName);
                this.convertPptToPdf(filePath, fileName);
                break;

            default:
                console.log("File type not supported");
        }
      },
      showDocViewer: function(name) {
          document.getElementsByClassName("doc-title")[0].innerHTML = name;
          (document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
          (document.getElementsByClassName("doc-container")[0]).style.visibility = "visible";
      },
      showPDFViewer: function(name) {
          (document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
          (document.getElementsByClassName("pdf-container")[0]).style.visibility = "visible";
          document.getElementsByClassName("pdf-title")[0].innerHTML = name;
      },
      showExcelViewer: function(name) {
          (document.getElementsByClassName("file-container")[0]).style.visibility = "hidden";
          (document.getElementsByClassName("excel-container")[0]).style.visibility = "visible";
          document.getElementsByClassName("excel-title")[0].innerHTML = name;
      },
      getFileStream: function(filePath, isPDF) {
          let ajax = new XMLHttpRequest();
          ajax.open("POST", hostUrl + "api/FileManager/GetDocument", true);
          ajax.setRequestHeader("content-type", "application/json");
          ajax.onreadystatechange = () => {
              if (ajax.readyState === 4) {
                  if (ajax.status === 200 || ajax.status === 304) {
                      if (!isPDF) {
                          this.$refs.docViewer.ej2Instance.documentEditor.open(ajax.responseText);
                      } else {
                          this.$refs.pdfViewer.load(ajax.responseText, null);
                      }
                  }
              }
          };
          ajax.send(JSON.stringify({ "FileName": filePath, "Action": (!isPDF ? "ImportFile" : "LoadPDF") }));
      },
      getBlob: function(filePath, fileName) {
          let request = new XMLHttpRequest();
          request.responseType = 'blob';
          request.onload = () => {
              let file = new File([request.response], fileName);
              this.$refs.excelViewer.open({ file: file });
          };
          request.open(
              'GET',
              hostUrl + 'api/FileManager/GetExcel' + '?FileName=' + filePath
          );
          request.send();
      },
      convertPptToPdf:function (filePath, fileName) {
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
                  this.showPDFViewer(fileName);
                  this.$refs.pdfViewer.load(pdfUrl, '');
              })
              .catch((error) => {
                  console.error('Error converting PPT to PDF:', error);
              });
      },
      closeView:function() {
        (document.getElementsByClassName("file-container")[0]).style.visibility = "visible";
        (document.getElementsByClassName("doc-container")[0]).style.visibility = "hidden";
        (document.getElementsByClassName("pdf-container")[0]).style.visibility = "hidden";
        (document.getElementsByClassName("excel-container")[0]).style.visibility = "hidden";
      }
    },
    //Injecting additional modules in FileManager
    provide: {
      filemanager: [DetailsView, NavigationPane, Toolbar]
    }
}
</script>
<style>
@import "../node_modules/@syncfusion/ej2-base/styles/material.css";
@import "../node_modules/@syncfusion/ej2-icons/styles/material.css";
@import "../node_modules/@syncfusion/ej2-inputs/styles/material.css";
@import "../node_modules/@syncfusion/ej2-popups/styles/material.css";
@import "../node_modules/@syncfusion/ej2-buttons/styles/material.css";
@import "../node_modules/@syncfusion/ej2-splitbuttons/styles/material.css";
@import "../node_modules/@syncfusion/ej2-navigations/styles/material.css";
@import "../node_modules/@syncfusion/ej2-layouts/styles/material.css";
@import "../node_modules/@syncfusion/ej2-grids/styles/material.css";
@import "../node_modules/@syncfusion/ej2-vue-filemanager/styles/material.css";
@import "../node_modules/@syncfusion/ej2-documenteditor/styles/material.css";
@import '../node_modules/@syncfusion/ej2-pdfviewer/styles/material.css';
@import '../node_modules/@syncfusion/ej2-spreadsheet/styles/material.css';
  .file-container {
    height:100%;
    width:100%;
  }
  .title-container {
    background-color: #eee;
    height: 22px;
  }
  .doc-container, .pdf-container, .excel-container {
    height:100%;
    width:100%;
    position: absolute;
    top: 0;
    visibility: hidden;
  }
  .close-button{
    float:right;
  }
  body,html{
    margin:0;
    height:100%;
    width:100%;
  }
</style>