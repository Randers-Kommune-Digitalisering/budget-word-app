import * as XLSX from 'xlsx';
import { sumArrays } from './utils.js';

// Constants
const dataSheetName = "Input";

const fileTypes = [
    "default",
    "expanded"
];

const dateTypes = [
    "31. marts",
    "31. maj",
    "30. september"
]

const defaultValues = [
    {column: "F", type: "select"},
    {column: "E", type: "select"},
    {column: "I", type: "select"},
    {column: "J", type: "select"},
    {columns: "KLM", type: "date_select"},
    {indexes: [4, 3], type: "diff"},
    {column: "O", type: "select"},
    {indexes: [5, 6], type: "sum"}
];

const defaultValuesExpanded = [
    {column: "D", type: "select"},
    {column: "C", type: "select"},
    {column: "G", type: "select"},
    {column: "H", type: "select"},
    {columns: "IJK", type: "date_select"},
    {indexes: [4, 3], type: "diff"},
    {column: "M", type: "select"},
    {indexes: [5, 6], type: "sum"}
];

// Variables
var globalData = null;

// Functions

export function readFile(file) {
    let sheets=["Input","Input - MTU","Input - BF - CT1","Input - SU - CT1","Input - EP - CT1"]
    const reader = new FileReader();
    reader.onload = function(evt) {
      if(evt.target.readyState != 2) return;
      if(evt.target.error) {
        throw Error("Kunne ikke læse filen. Prøv igen.");
      }
      try {
        let excel_file = XLSX.read(evt.target.result);
        globalData=[]
        for (let sheet in sheets) {
          console.log("Sheet", sheet)
          let data_sheet = excel_file.Sheets[sheets[sheet]];
          console.log("data_sheet", data_sheet)
          if (data_sheet) {
              globalData.push(data_sheet);
              console.log("GlobalData", globalData)
          } else {
            throw("Kunne ikke finde Input-arket. Prøv igen.");
          }
        }
      } catch (error) {
        console.log(error);
        throw("Kunne ikke læse filen. Prøv igen.");
      };
    }
    reader.readAsArrayBuffer(file);
}

export function generateTable(columns, rows, withData, dateType, fileType, sheet=0) {
    let table = [columns]
    for (var i in rows){
      if(Array.isArray(rows[i]) && rows[i].length > 1) {
        let id = null
        if(rows[i].length === 2) id = rows[i][1]
        else id = rows[i].slice(1)
        if(id && withData){
          let first_column = rows[i][0]
          let data = getRowData(id, dateTypes.indexOf(dateType), globalData[sheet], fileTypes.indexOf(fileType))
          table.push([first_column, ...data])
        } else {
            let tmp_row = [rows[i][0]]
            rows[i] = tmp_row[0]
            for(var j = 1; j <= columns.length-1; j++) {
              tmp_row.push(null)
            }
            table.push(tmp_row)
        }
      } else {
        var tmp_row=[rows[i]]
        for(var j = 1; j <= columns.length-1; j++) {
          tmp_row.push(null)
        }
        table.push(tmp_row)
      }
    }
    return table
}

function getRowData(id, date, sheet, fileType) {

    if(!sheet) throw Error("No data sheet")
    if(!id) throw Error("No id")

    let selectedValues = null

    if(fileType === fileTypes[1]) selectedValues = defaultValuesExpanded
    else selectedValues = defaultValues

    function getRowById(expaned, sheet, id) {
        if(expaned) return Object.entries(sheet).filter(([key, value]) => (key.includes("A"))).find(arr => String(arr[1].v).includes(id))
        else return Object.entries(sheet).filter(([key, value]) => (key.includes("A") || key.includes("C"))).find(arr => String(arr[1].v).includes(id))
    }

    let rows = []

    if(Array.isArray(id)) {
        id.forEach(id => {
            let row = (getRowById(false, sheet, id) ? getRowById(false, sheet, id)[0].slice(1) : undefined)
            rows.push(row)
        })
    } else {
        let row = (getRowById(false, sheet, id) ? getRowById(false, sheet, id)[0].slice(1) : undefined)
        rows.push(row)
    }

    if(rows){
        let totalValues = []
        rows.forEach(row => {
            if(row) { 
                let values = []
                selectedValues.forEach(key => {
                    if(key.column && key.type === "select") {
                        if(!sheet[`${key.column}${row}`]) values.push(0)
                        else values.push(sheet[`${key.column}${row}`].v)
                    } else if(key.columns && key.type === "date_select") {
                        if(!sheet[`${key.columns[date]}${row}`]) values.push(0)
                        else values.push(sheet[`${key.columns[date]}${row}`].v)
                    } else if(key.indexes) {
                        if (key.type === "diff") {
                            let runningTotal = undefined
                            key.indexes.forEach(index => {
                                if(runningTotal === undefined) runningTotal = values[index]
                                else runningTotal -= values[index]
                            })
                            values.push(runningTotal)
                        } else if (key.type === "sum") {
                            let runningTotal = 0
                            key.indexes.forEach(index => {
                                runningTotal += values[index]
                            })
                            values.push(runningTotal)
                    } else {
                        console.log("unknown type with indexes", key)
                    }
                    } else {
                    console.log("unknown type", key)
                    }
                });
                //values = values.map(value => (Math.round(value * 10) / 10).toFixed(1))   Flyttes da det giver afrundingsfejl
                totalValues.push(values)
            } else throw Error("Data error: row for id '" + id + "' not found")
        });
        let res = null
        if(totalValues.length > 1){
            res = sumArrays(...totalValues.map( subarray => subarray.map( (el) => parseFloat(el)))) // back to floats and sum them
            //res = res.map(value => (Math.round(value * 10) / 10).toFixed(1)) // round again
        } else {
            res = totalValues[0]
        }
        return res
    } else throw Error("Data error")
}

function hasData(sheet, ids, fileType) {
  let rows = []
  ids.forEach(id => {
    // fileType 1 is expanded
    if(fileType === fileTypes[1]) rows.push(Object.entries(sheet).filter(([key, value]) => (key.includes("A"))).find(arr => String(arr[1].v).includes(id)))
    else rows.push(Object.entries(sheet).filter(([key, value]) => (key.includes("A") || key.includes("C"))).find(arr => String(arr[1].v).includes(id)))
  })

  return rows.some(item => item === undefined) ? false : true

}