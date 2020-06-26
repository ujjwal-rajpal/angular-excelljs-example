import { Component } from '@angular/core';
// import { data }
import { ExcellService } from "./service/excell.service"
import * as data from './service/data'
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'angular-excelljs-example';

  constructor( private  _excellService: ExcellService ){}

  downloadExcell(  ){
    let bookedHeader = [], bookedData= [], transitData = [], transitHeader = [], deleveredData = [], deleveredHeader = []
    let excellData = JSON.parse( data.data.response.scriptResult)
    
    let tableObjects = excellData.tableObjects
    let dataArray = []
    tableObjects.forEach(element => {
      
      let data = element.data.split(`EMPTY`)
      let excellHeaders: [] = (((data[0].replace(/['"]+/g, '')).split(",")))
      let excellData: []= (data[1].replace(/['"]+/g, '')).split(",")
      
      if(element.name === 'Booked'){
        bookedHeader = ["SN", "Trip#", "Truck#", "Driver Name", "Phone Number", "Transporter", "BL Number", "Cargo Name", "Customer PO Lot # Invoice #", "Destination", "Gross Tonnage", "Nett Tonnage", "Packages", "Current Location", "Status", "ETA"]
        var i,j,  tempaDataArray=[]
        for( i=1, j=1; i<= excellData.length; i++){
          
          tempaDataArray.push(excellData[i-1])
          if(i === j*(excellHeaders.length - 1)){

            let i =[0,2,3,8,21,7,16,14,41, 9, 10, 11, 12, 40, 46, 22]
            let newTempArray = []
            i.forEach(element=>{  
              
              if(element === 0 || element === 11 || element === 12  || element === 10 ){
                if(tempaDataArray[element])
                tempaDataArray[element] = parseInt(tempaDataArray[element])
              }
              
              //
              newTempArray.push(tempaDataArray[element])
            })
            bookedData.push(newTempArray)
            tempaDataArray=[]
            j++
          }
        }
       
      
      } else if(element.name === 'InTransitData'){
        transitHeader = ["SN", "Trip#", "Truck#", "Driver Name", "Phone Number", "Transporter", "BL Number", "Cargo Name", "Customer PO Lot # Invoice #", "Destination", "Gross Tonnage", "Nett Tonnage", "Packages", "Current Location", "Status", "Total Trip Days", "Dispatch Date", "Arrive Load Site", "Loading Date", "Leave Load Site", "Dwell", "Arrive Tunduma", "Leave Tunduma", "Dwell", "Arrive Kbp (Sakania)", "Leave Kbp (Sakania)", "Dwell", "Arive Offload Site", "Offload Date", "Dwell"]
        var i,j,  tempaDataArray=[]
        for( i=1, j=1; i<= excellData.length; i++){
          
          tempaDataArray.push(excellData[i-1])
          if(i === j*(excellHeaders.length - 1)){

            let i =[0,2,3,8,,21,7,16,14,41, 9, 10, 11, 12, 40, 46, 39, 23, 24, 25, 26, 27, 29, 30, 31, 32, 33, 34, 48, 37, 38]
            let newTempArray = []
            
            i.forEach(element=>{  
              
              if(element === 0 || element === 11 || element === 12  || element === 10 || element === 23 || element === 27 || element === 31 || element === 34 || element === 38 )
              if(tempaDataArray[element])
              tempaDataArray[element] = parseInt(tempaDataArray[element])
              newTempArray.push(tempaDataArray[element])
            })
            transitData.push(newTempArray)
            tempaDataArray=[]
            j++
          }
        }

      }else if(element.name === 'DeliveredData'){
        deleveredHeader = ["SN", "Trip#", "Truck#", "Driver Name", "Phone Number", "Transporter", "BL Number", "Cargo Name", "Customer PO Lot # Invoice #", "Destination", "Gross Tonnage", "Nett Tonnage", "Packages", "Current Location", "Status", "Total Trip Days", "Dispatch Date", "Arrive Load Site", "Loading Date", "Leave Load Site", "Dwell", "Arrive Tunduma", "Leave Tunduma", "Dwell", "Arrive Kbp (Sakania)", "Leave Kbp (Sakania)", "Dwell", "Arive Offload Site", "Offload Date", "Dwell"]
        var i,j,  tempaDataArray=[]
        for( i=1, j=1; i<= excellData.length; i++){
          
          tempaDataArray.push(excellData[i-1])
          if(i === j*(excellHeaders.length - 1)){

            let i =[0,2,3,8,,21,7,16,14,41, 9, 10, 11, 12, 40, 46, 39, 23, 24, 25, 26, 27, 29, 30, 31, 32, 33, 34, 48, 37, 38]
            let newTempArray = []
            i.forEach(element=>{  
              if(element === 0 || element === 11 || element === 12  || element === 10 || element === 23 || element === 27 || element === 31 || element === 34 || element === 38 )
              if(tempaDataArray[element])
              tempaDataArray[element] = parseInt(tempaDataArray[element] )
              newTempArray.push(tempaDataArray[element])
            })
                deleveredData.push(newTempArray)
            tempaDataArray=[]
            j++
          }
        }
      }
      
    });
    
    

    let reportData = {
      bookedHeader: bookedHeader,
      bookedData: bookedData,
      transitHeader: transitHeader,
      transitData: transitData,
      deleveredHeader: deleveredHeader,
      deleveredData: deleveredData
    }
    // console.log( Object.keys(empPerformance[0]))
    this._excellService.generateExcell(reportData);
  }
}
