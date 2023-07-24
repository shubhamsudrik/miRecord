import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-file-upload',
  templateUrl: './file-upload.component.html',
  styleUrls: ['./file-upload.component.css']
})
export class FileUploadComponent {


   data : [][];
   ngOninit(): void {

   }

   onFileChange(evt: any){
     const target: DataTransfer =<DataTransfer>(evt.target);

      console.log(evt.target.files);

      const reader: FileReader = new FileReader();

      reader.onload = (a : any) =>{
         const bstr: string = a.target.result;
         const wb : XLSX.WorkBook =XLSX.read(bstr, {type: 'binary' });
         const wsname : string =wb.SheetNames[0];
         const ws: XLSX.WorkSheet =wb.Sheets[wsname];

         console.log(ws);


         this.data = (XLSX.utils.sheet_to_json(ws,{header:1 }));
         console.log(this.data);
      };

      reader.readAsBinaryString(target.files[0]);
   }
}
