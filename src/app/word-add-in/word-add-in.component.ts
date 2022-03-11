import { Component, OnInit, ViewChild } from '@angular/core';

import { UtilityService } from '../service/utility.service';

declare var Word:any, fabric;

@Component({
  selector: 'app-word-add-in',
  templateUrl: './word-add-in.component.html',
  styleUrls: ['./word-add-in.component.css']
})
export class WordAddInComponent implements OnInit {

  dialogComponent: any;
  overlayComponent: any;
  buttonContent: string;

  @ViewChild('dialogue',{ static : false }) dialogue;
  @ViewChild('overlay',{ static : false }) overlay;
  
  constructor(private utilityService:UtilityService) { }

  ngOnInit() {
  }

  ngAfterViewInit(){
    this.dialogComponent = new fabric['Dialog'](this.dialogue.nativeElement);
    this.overlayComponent = new fabric['Overlay'](this.overlay.nativeElement);
  }
  
  openDialogue(){
    this.dialogComponent.open();
    this.overlayComponent.show();
  }

  repeatTableContent(){
    this.insertTableContent(this.buttonContent);
    this.overlayComponent.hide();
  }

  dontRepeatTableContent(){
    this.insertContent(this.buttonContent);
    this.overlayComponent.hide();
  }

  insertText(content){
    this.buttonContent = content;
    Word.run((context) => {
         var range = context.document.getSelection();
         var tableCell = range.parentTableCellOrNullObject;
         context.load(tableCell);
         return context.sync().then(()=>{
             if(tableCell.isNull){ 
                this.insertContent(this.buttonContent);
             }
             else{
                this.openDialogue();
             }   
         });
    });
  }

  insertContent(text) {
    Word.run( (context) => {
       var range = context.document.getSelection();
       range.insertText(text,"Replace");
       return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
  }

  insertTableContent(text){
    Word.run((context)=>{

      let range, tableCell, table;
      range = context.document.getSelection();
      tableCell = range.parentTableCellOrNullObject;
      table = range.parentTableOrNullObject;
      context.load(tableCell);
      context.load(table);

      return context.sync().then(()=>{
        let columnIndex = tableCell.cellIndex ,i:number, tag:number = 0, rowIndex, rowCount;
        rowIndex = tableCell.rowIndex;
        rowCount = table.rowCount;
        for(i = rowIndex; i < rowCount; i++,tag++){
           let cell = table.getCellOrNullObject(i,columnIndex);
           let content = `${text}_${tag}`;
           cell.value = content;
        }
      });

    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  uploadFile(){
    Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
            if (result.status == Office.AsyncResultStatus.Succeeded) {
                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };
                console.log("Getting file of " + myFile.size + " bytes");
                this.getSlice(state);
            }
            else {
                console.log(result.status);
            }
        });
  }

  getSlice(state) {
    let that = this;
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            console.log("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            that.sendSlice(result.value, state);
        }
        else {
            console.log(result.status);
        }
    });
  }

  sendSlice(slice, state) {
    let data = slice.data, byteArray, fileUrl, splitUrl, fileName, file;
    fileName = 'Document.docx';
    byteArray = Uint8Array.from(data);
    fileUrl = Office.context.document.url;
    if(fileUrl){
      splitUrl = fileUrl.split("/");
      fileName = splitUrl[splitUrl.length-1];
    }
    file = new File([byteArray],fileName, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const formData = new FormData();
    formData.set('file', file);
    this.utilityService.uploadDocToCommusoft({}, formData).subscribe(
      response=>{
         console.log(response);
      },
      error=>{
         console.log(error);
    });

    // var url = window.URL.createObjectURL(file);
    // var a = document.createElement('a');
    // a.style.display = 'none';
    // a.href = url;
    // a.download = fileName;
    // document.body.appendChild(a);
    // a.click();
    // setTimeout(function() {
    // document.body.removeChild(a);
    // window.URL.revokeObjectURL(url);
    // }, 100);

    this.closeFile(state);
  }

  closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {
        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status == "succeeded") {
            console.log("File closed.");
        }
        else {
            console.log("File couldn't be closed.");
        }
    });
}

}
