import { Component, Input, OnInit } from '@angular/core';
import { CoreBase, IMIRequest, IMIResponse, IUserContext, MIRecord } from '@infor-up/m3-odin';
import { MIService, UserService } from '@infor-up/m3-odin-angular';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent extends CoreBase implements OnInit {
  userContext = {} as IUserContext;

  lineitems: MIRecord[];
  transData: MIRecord[];
  transDetails: MIRecord[];
  FLNM: string;
  trtpData = 'I';
  selectedMINM: string = '';
  selectedTransaction: string = '';
  selectedTransactionData: any = null;

  constructor(private miService: MIService, private userService: UserService) {
    super('AppComponent');
  }
  ngOnInit() {
    this.userService.getUserContext().subscribe(
      (userContext: IUserContext) => {
        this.callAllFunctions();
      },
      (error) => {
        console.log('Failed to get user context:', error);
      }
    );
  }

  callAllFunctions() {
    this.MRS001MI_LstPrograms();
  }

  MRS001MI_LstPrograms() {
    const inputRecord = new MIRecord();
    const request: IMIRequest = {
      program: 'MRS001MI',
      transaction: 'LstPrograms',
      record: inputRecord,
      outputFields: ['MINM'],
      maxReturnedRecords: 0,
    };
    console.log(request);
    this.miService.execute(request).subscribe(
      (response: IMIResponse) => {
        if (!response.hasError()) {
          this.lineitems = response.items as MIRecord[];
        }
        console.log(this.lineitems);

      },
      (error) => {
        console.log('Error executing MRS001MI LstPrograms:', error);
      }
    );
  }

  MRS001MI_LstTransaction(minmData: string) {
    const inputRecord = new MIRecord();
    inputRecord.setNumber('MINM', minmData);
    const request: IMIRequest = {
      program: 'MRS001MI',
      transaction: 'LstTransactions',
      record: inputRecord,
      outputFields: ['TRNM'],

      maxReturnedRecords: 0,
    };

    this.miService.execute(request).subscribe(
      (response: IMIResponse) => {
        if (!response.hasError()) {
          this.transData = response.items as MIRecord[];
          console.log(this.transData);
        }
      },
      (error) => {
        console.log('Error executing MRS001MI LstTransactions:', error);
      }
    );
  }

  MRS001MI_LstFields(trnmData: string, minmData:string ,trtpData:string) {
    const inputRecord = new MIRecord();
    inputRecord.setString('TRNM', trnmData);
    inputRecord.setString('MINM', minmData);
    inputRecord.setString('TRTP', trtpData);

    const request: IMIRequest = {
      program: 'MRS001MI',
      transaction: 'LstFields',
      record: inputRecord,
      outputFields: ['FLNM', 'FLDS','TXT1', 'FRPO','TOPO','LENG','TYPE','MAND'],
    };

    this.miService.execute(request).subscribe(
      (response: IMIResponse) => {
        if (!response.hasError()) {
          this.transDetails = response.items as MIRecord[];
        }
        console.log(this.transDetails);

      },
      (error) => {
        console.log('Error executing MRS001MI LstFields:', error);
      }
    );
  }

  onClickMIRecord(minm: string) {
    this.selectedMINM = minm;
    this.MRS001MI_LstTransaction(minm);
    console.log(minm);
    this.selectedTransaction = '';
    this.selectedTransactionData = null;
  }


  onClickTransaction(trnm: string) {
    this.selectedTransaction = trnm;
    this.MRS001MI_LstFields(trnm, this.selectedMINM, this.trtpData);

  }
ngoninit(){
  console.log(this.transDetails);
}
  // ***************************************************************************************************************
//i want to download this mi data in excel format
  fileName= 'ExcelSheet.xlsx';

  exportexcel(): void {
    const element = document.getElementById('excel-table'); //excel-table = Id
    const rows = element.getElementsByTagName('tr');
    const columnData: string[] = [];

  console.log(rows);
  console.log(this.transData);

    // Iterate over the rows and extract the first column data
    for (let i = 0; i < rows.length; i++) {
      const columns = rows[i].getElementsByTagName('td');
      console.log(columns);
      if (columns.length > 0) {
        const cellValue = columns[0].innerText;
        const cellValue1 = columns[1].innerText; //[column_index_value]
        console.log(cellValue)
        columnData.push(cellValue);
        columnData.push(cellValue1)
      }
    }

    // Create a new workbook and worksheet
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet([columnData]);

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // Save the workbook as Excel file
    XLSX.writeFile(wb, this.fileName);
  }



}
