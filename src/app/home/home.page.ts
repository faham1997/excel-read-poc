import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-home',
  templateUrl: 'home.page.html',
  styleUrls: ['home.page.scss'],
})
export class HomePage implements OnInit {
  datas: any[];
  constructor() {}

  ngOnInit(): void {}

  public onFileSelected = (event: any) => {
    let data: any;
    const file: File = event.target.files[0];
    // the method for reading multiple lines
    // this.readFile(file);
    const fileReader = new FileReader();

    fileReader.onload = (e: any) => {
      const arrayBuffer = e.target.result;
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const worksheetName = workbook.SheetNames[4];
      const worksheet = workbook.Sheets[worksheetName];
      var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      console.log(jsonData);
      // this.generateExcelFile(this.datas);
    };

    fileReader.readAsArrayBuffer(file);
  };

  //for reading multiple sheets
  public readFile = (file: File) => {
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const data: Uint8Array = new Uint8Array(e.target.result);
      const workbook: XLSX.WorkBook = XLSX.read(data, { type: 'array' });

      workbook.SheetNames.forEach((page: string) => {
        const worksheet: XLSX.WorkSheet = workbook.Sheets[page];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          raw: false,
          header: 1,
        });
        console.log(`Page Name: ${page}`);
        console.log(jsonData);
      });
    };
    reader.readAsArrayBuffer(file);
  };

  generateExcelFile(data): void {
    this.datas = data;
    console.log(this.datas);

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const excelBuffer = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });

    this.saveAsExcelFile(excelBuffer, 'sample');
  }

  saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    // Use FileSaver.js or similar library to save the file
    saveAs(data, fileName + '_export_' + new Date().getTime() + '.xlsx');
  }
}
