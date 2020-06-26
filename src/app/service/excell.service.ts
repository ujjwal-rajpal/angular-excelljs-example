import { Injectable } from '@angular/core';

// import { Workbook } from 'exceljs';
    import * as logo from './logo'
// import * as Excel from 'exceljs';
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as ExcelProper from "exceljs";

import * as fileSaver from 'file-saver';




@Injectable({
  providedIn: 'root'
})
export class ExcellService {

  constructor() { }


  generateExcell(excelldata){
    console.log(excelldata)
    const companyName = 'Blue Gorilla Corporation Limited'
    const companyAddressLine1 = `4th Floor, Ebene Skies, Rue de L'Institut`
    const companyAddressLine2 = `Ebene, Port Louis 8817 Mauritius`
    // let workbook = new Excel.Workbook();
    let workbook: ExcelProper.Workbook = new Excel.Workbook();
    let worksheet= workbook.addWorksheet('Pre-Alerted',  {views: [{showGridLines: false}]} );
    let worksheet2= workbook.addWorksheet('Tracking',  {views: [{showGridLines: false}]} );
    let worksheet3= workbook.addWorksheet('Delivered',  {views: [{showGridLines: false}]} );
    // let worksheet2= workbook.addWorksheet('carData',  {views: [{showGridLines: false}]});

    // add images to all three workbooks
    let logos = workbook.addImage({
      base64: logo.logoBase64,
      extension: 'png',
    });
    // for preAlert
    worksheet.mergeCells('A1:c4')
    worksheet.addImage(logos,
      {
        tl: { col: 0.5, row: 0.5 },
        br: { col: 2.5, row: 3.5 }
      }
      )

      // for tracking
      worksheet2.mergeCells('A1:c4')
    worksheet2.addImage(logos,
      {
        tl: { col: 0.5, row: 0.5 },
        br: { col: 2.5, row: 3.5 }
      }
      )

      // for delivered
      worksheet3.mergeCells('A1:c4')
    worksheet3.addImage(logos,
      {
        tl: { col: 0.5, row: 0.5 },
        br: { col: 2.5, row: 3.5 }
      }
      )

    // insert at A5
    //  pre-Alert
    let comapnyNameCell = worksheet.getCell('A5');
    comapnyNameCell.value = companyName
    // tracking
    comapnyNameCell = worksheet2.getCell('A5');
    comapnyNameCell.value = companyName
    //delivered
    comapnyNameCell = worksheet3.getCell('A5');
    comapnyNameCell = worksheet3.getCell('A5');

    // insert at A6
    let companyAddressLine1Cell =  worksheet.getCell('A6')
    companyAddressLine1Cell.value = companyAddressLine1
    //tracking
    companyAddressLine1Cell =  worksheet2.getCell('A6')
    companyAddressLine1Cell.value = companyAddressLine1
    //delivered
    companyAddressLine1Cell =  worksheet3.getCell('A6')
    companyAddressLine1Cell.value = companyAddressLine1


    // insert At A7 pre-Alert
    let companyAddressLine2Cell = worksheet.getCell('A7')
    companyAddressLine2Cell.value = companyAddressLine2
    // tracking
    companyAddressLine2Cell = worksheet2.getCell('A7')
    companyAddressLine2Cell.value = companyAddressLine2
    // delivered
    companyAddressLine2Cell = worksheet3.getCell('A7')
    companyAddressLine2Cell.value = companyAddressLine2

    // Right section pre-Alert
    // worksheet.mergeCells('z1:AD1') 
    let infoline1Cell = worksheet.getCell('AD1')
    infoline1Cell.value = "[This is system generated report]"
    infoline1Cell.alignment = { vertical: 'middle', horizontal: 'right' }
    infoline1Cell.font = {
      name: 'Calibri',
      size: 8
    }

    // tracking
    infoline1Cell = worksheet2.getCell('AD1')
    infoline1Cell.value = "[This is system generated report]"
    infoline1Cell.alignment = { vertical: 'middle', horizontal: 'right' }
    infoline1Cell.font = {
      name: 'Calibri',
      size: 8
    }
    
    // delivered
    infoline1Cell = worksheet3.getCell('AD1')
    infoline1Cell.value = "[This is system generated report]"
    infoline1Cell.alignment = { vertical: 'middle', horizontal: 'right' }
    infoline1Cell.font = {
      name: 'Calibri',
      size: 8
    }

    // preAlert
    let infoline2Cell =  worksheet.getCell('AD2')
    infoline2Cell.value = "[Please note that the variations may occurs as we are tracking moving assets]"
    infoline2Cell.alignment = { vertical: 'middle', horizontal: 'right' }
    infoline2Cell.font = {
      name: 'Calibri',
      size: 8
    }

    // transit 
    infoline2Cell =  worksheet2.getCell('AD2')
    infoline2Cell.value = "[Please note that the variations may occurs as we are tracking moving assets]"
    infoline2Cell.alignment = { vertical: 'middle', horizontal: 'right' }
    infoline2Cell.font = {
      name: 'Calibri',
      size: 8
    }

    // delivered
    infoline2Cell =  worksheet3.getCell('AD2')
    infoline2Cell.value = "[Please note that the variations may occurs as we are tracking moving assets]"
    infoline2Cell.alignment = { vertical: 'middle', horizontal: 'right' }
    infoline2Cell.font = {
      name: 'Calibri',
      size: 8
    }

    // preAllert
    let emergencyContact = worksheet.getCell('AD4')
    emergencyContact.value = "[ Emergency Contact: +230 404 8034]"
    emergencyContact.alignment = { vertical: 'middle', horizontal: 'right' }
    emergencyContact.font = {
      name: 'Calibri',
      size: 8,
      bold: true
    }

    // tracking
    emergencyContact = worksheet.getCell('AD4')
    emergencyContact.value = "[ Emergency Contact: +230 404 8034]"
    emergencyContact.alignment = { vertical: 'middle', horizontal: 'right' }
    emergencyContact.font = {
      name: 'Calibri',
      size: 8,
      bold: true
    }
    // delivered
    emergencyContact = worksheet.getCell('AD4')
    emergencyContact.value = "[ Emergency Contact: +230 404 8034]"
    emergencyContact.alignment = { vertical: 'middle', horizontal: 'right' }
    emergencyContact.font = {
      name: 'Calibri',
      size: 8,
      bold: true
    }

    // empty row pre-Allert
    worksheet.addRow([]);
    // Intransit
    worksheet2.addRow([]);
    // delivered
    worksheet3.addRow([]);
    // pre-Allert
    worksheet.mergeCells('A9:P9')
    // delivered
    worksheet2.mergeCells('A9:AD9')
    // Intransit
    worksheet3.mergeCells('A9:AD9')

    // add heading pre-Allert
    let headingCell = worksheet.getCell('A9')
    
    headingCell.value = "Booked Trips"
    headingCell.font = {
      name: 'Calibri',
      size: 12,
      bold: true

    }
    headingCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'deebf7' },
      bgColor: { argb: '' }
    }
    headingCell.alignment = { vertical: 'middle', horizontal: 'center' }

    // transit
    headingCell = worksheet2.getCell('A9')
    
    headingCell.value = "Going Load Tracking Report"
    headingCell.font = {
      name: 'Calibri',
      size: 12,
      bold: true

    }
    headingCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'deebf7' },
      bgColor: { argb: '' }
    }
    headingCell.alignment = { vertical: 'middle', horizontal: 'center' }

    // delivered
    headingCell = worksheet3.getCell('A9')
    
    headingCell.value = "Delivered Lots"
    headingCell.font = {
      name: 'Calibri',
      size: 12,
      bold: true

    }
    headingCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'deebf7' },
      bgColor: { argb: '' }
    }
    headingCell.alignment = { vertical: 'middle', horizontal: 'center' }



    // change row height pre-Alert
    let row = worksheet.lastRow;
    row.height = 35
    //In-transit
    row = worksheet2.lastRow;
    row.height = 35
    //deliverd
    row = worksheet3.lastRow;
    row.height = 35

    //Adding Header Row Pre-Alert
    let headerRow = worksheet.addRow(excelldata. bookedHeader);
    headerRow.eachCell((cell, number) => {
      cell.font = {
        bold: true,
        color: { argb: '000000' },
        size: 12
      }
      let column = worksheet.getColumn(number)
      switch(number){
        case 1:
           column.width = 5
        break;
        case 2:
        case 3:  
        case 7:
        case 11:
        case 12:
        case 13:
        case 17:
        case 18:
        case 19:
        case 20:
        case 22:
        case 23:
        case 25:
        case 26:
        case 28:
        case 29:
        case 16:
           column.width = 15
        break
        case 4:
        case 8:
          column.width = 35
        break
        case 5:
        case 9: 
        case 10:
          column.width = 20
        break
        case 6 : column.width = 45
        break
        case 14:
        case 15: column.width = 40
        break
        
        case 21:
        case 24:
        case 27:
        case 30:column.width=7
        break
      }
    })

    // In-transit
    headerRow = worksheet2.addRow(excelldata.transitHeader);
    headerRow.eachCell((cell, number) => {
      cell.font = {
        bold: true,
        color: { argb: '000000' },
        size: 12
      }
      let column = worksheet2.getColumn(number)
      switch(number){
        case 1:
           column.width = 5
        break;
        case 2:
        case 3:  
        case 7:
        case 11:
        case 12:
        case 13:
        case 17:
        case 18:
        case 19:
        case 20:
        case 22:
        case 23:
        case 25:
        case 26:
        case 28:
        case 29:
           column.width = 15
        break
        case 4:
        case 8:
          column.width = 35
        break
        case 5:
        case 9: 
        case 10:
          column.width = 20
        break
        case 6 : column.width = 45
        break
        case 14:
        case 15: column.width = 40
        break
        case 16:
        case 21:
        case 24:
        case 27:
        case 30:column.width=7
        break
      }
    })
    
    // delivered
    headerRow = worksheet3.addRow(excelldata.deleveredHeader);
    headerRow.eachCell((cell, number) => {
      cell.font = {
        bold: true,
        color: { argb: '000000' },
        size: 12
      }
      let column = worksheet3.getColumn(number)
      switch(number){
        case 1:
           column.width = 5
        break;
        case 2:
        case 3:  
        case 7:
        case 11:
        case 12:
        case 13:
        case 17:
        case 18:
        case 19:
        case 20:
        case 22:
        case 23:
        case 25:
        case 26:
        case 28:
        case 29:
           column.width = 15
        break
        case 4:
        case 8:
          column.width = 35
        break
        case 5:
        case 9: 
        case 10:
          column.width = 20
        break
        case 6 : column.width = 45
        break
        case 14:
        case 15: column.width = 40
        break
        case 16:
        case 21:
        case 24:
        case 27:
        case 30:column.width=7
        break
      }
    })
    





    // 
    // rotate cell and alignment
    let cells = ['A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10', 'J10', 'K10', 'L10', 'M10', 'N10', 'O10', 'P10', 'Q10', 'R10','S10', 'T10', 'U10', 'V10', 'W10', 'X10', 'Y10', 'Z10', 'AA10', 'AB10', 'AC10', 'AD10']
    cells.forEach(element => {
    
      if(element === 'P10' || element === 'U10' || element === 'X10' || element === 'AA10' || element === 'AD10'){
        worksheet.getCell(element).alignment = {  horizontal:"left", vertical: 'middle', wrapText: true , textRotation: 90 };
        worksheet2.getCell(element).alignment = {  horizontal:"left", vertical: 'middle', wrapText: true , textRotation: 90 };
        worksheet3.getCell(element).alignment = {  horizontal:"left", vertical: 'middle', wrapText: true , textRotation: 90 };
      }
      else if( element === 'A10' || element === 'Q10' || element === 'R10' || element === 'S10' || element === 'T10' || element === 'V10' || element === 'w10' || element === 'Y10' || element === 'Z10' || element === 'AB10' || element === 'AC10'){
        worksheet.getCell(element).alignment = {  horizontal:"center", vertical: 'middle', wrapText: true };
        worksheet2.getCell(element).alignment = {  horizontal:"center", vertical: 'middle', wrapText: true };
        worksheet3.getCell(element).alignment = {  horizontal:"center", vertical: 'middle', wrapText: true };
      }
      else{
        worksheet.getCell(element).alignment = {  horizontal:"left", vertical: 'middle', wrapText: true };
        worksheet2.getCell(element).alignment = {  horizontal:"left", vertical: 'middle', wrapText: true };
        worksheet3.getCell(element).alignment = {  horizontal:"left", vertical: 'middle', wrapText: true };
      }      
    });

    // rerotale P10 of worksheet 1
    worksheet.getCell('P10').alignment = {  horizontal:"center", vertical: 'middle', wrapText: true };

    let header = worksheet.lastRow;
    header.height = 70
    header = worksheet2.lastRow;
    header.height = 70
    header = worksheet3.lastRow;
    header.height = 70

    // Adding Data with Conditional Formatting #in-transit
    excelldata.bookedData.forEach(d => {
      let row = worksheet.addRow(d);

      let currentLocation = row.getCell(14);
      let status = row.getCell(15)
    
      currentLocation.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }

      status.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      
    
    });

    //tracking
    excelldata.transitData.forEach(d => {
      let row = worksheet2.addRow(d);

      let currentLocation = row.getCell(14);
      let status = row.getCell(15)
      let loadingDate = row.getCell(19)
      let dwell = row.getCell(21)
      let dwell2 = row.getCell(24)
      let dwell3 = row.getCell(24)
      let dwell4 = row.getCell(24)
      currentLocation.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }

      status.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      
      loadingDate.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell2.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell3.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell4.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
    });

    //tracking
    excelldata.deleveredData.forEach(d => {
      let row = worksheet3.addRow(d);

      let currentLocation = row.getCell(14);
      let status = row.getCell(15)
      let loadingDate = row.getCell(19)
      let dwell = row.getCell(21)
      let dwell2 = row.getCell(24)
      let dwell3 = row.getCell(24)
      let dwell4 = row.getCell(24)
      currentLocation.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }

      status.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      
      loadingDate.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell2.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell3.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
      dwell4.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'bdd7ee' }
      }
    });



//     // Add row with current date
//     // let subtitleRow = worksheet.addRow(['Date: ' + this.datePipe.transform(new Date(), 'medium')])
// // add image


workbook.xlsx.writeBuffer().then(data=>{
  let date = new Date()
  let blob = new Blob([data], { type: 'application/vnd.openXmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'})
  fileSaver.saveAs(blob,`LEO_SubContractorTrackingFS_${date.getFullYear()}-${date.getMonth()+1}-${date.getDate()}.xlsx`)
})
  }
}
