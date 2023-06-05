import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-home',
  templateUrl: 'home.page.html',
  styleUrls: ['home.page.scss'],
})
export class HomePage {
  constructor() {}

  public onFileSelected = (event: any) => {
    const file: File = event.target.files[0];
    this.readFile(file);
  };

  public readFile = (file: File) => {
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const data: Uint8Array = new Uint8Array(e.target.result);
      const workbook: XLSX.WorkBook = XLSX.read(data, { type: 'array' });

      workbook.SheetNames.forEach((page: string) => {
        const worksheet: XLSX.WorkSheet = workbook.Sheets[page];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });
        console.log(`Page Name: ${page}`);
        console.log(jsonData);
      });
    };
    reader.readAsArrayBuffer(file);
  };
}
