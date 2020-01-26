import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})

export class AppComponent {
  title = 'ng-read-excel';
  result: any = {};
  sheets = [];
  sheet = '';
  constructor() {

  }

  fileChangeEvent(evt) {
    this.sheets = [];
    this.result = {};
    const target: DataTransfer = (evt.target) as DataTransfer;
    if (target.files.length !== 1) { throw new Error('Cannot use multiple files'); }
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary', cellDates: true});
      wb.SheetNames.forEach(element => {
        this.sheets.push(element);
        const ws: XLSX.WorkSheet = wb.Sheets[element];
        this.result[element] = XLSX.utils.sheet_to_json(ws, {header: 1, raw: false});
      });
      this.sheet = wb.SheetNames[0];
    };
    reader.readAsBinaryString(target.files[0]);
  }

  onClick(sheet: string) {
    this.sheet = sheet;
  }
}
