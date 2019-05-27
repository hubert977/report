const express = require('express');
const app = express();
const mysql = require('mysql');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();
class excel 
{
    constructor()
    {
        this.configure = {
            host: '10.10.2.241',
            user: 'sa',
            password: 'Radwag99',
            database: 'pue71'
        }
    }
    initialize()
    {
        console.log('asdfdasf');
        let correctArray = [];
        const connection = mysql.createConnection(this.configure);
            connection.query('SELECT id_comparator from weighting ORDER BY id_comparator DESC LIMIT 220', (error, results, fields) => {
            if (error) throw error;
                let results_json = JSON.stringify(results);
                let resultjson = JSON.parse(results_json);
                const that = this;
                new Promise(function(resolve, reject) {
                    let arrayvalue = [];
                    let arrayData = [];
                    for(let i=0; i<resultjson.length; i++)
                    {
                         arrayvalue.push(resultjson[i].id_comparator);
                         arrayData.push(resultjson[i]);
                         correctArray = [...new Set(arrayvalue)];
                    }
                    that.cyclesiterable(correctArray);
                    console.log(correctArray);
                  })
        });
    }
    cyclesiterable(correctArray)
    {
        const connection = mysql.createConnection(this.configure);
        for(let i=0; i<correctArray.length; i++)
        {

            connection.query(`SELECT comparatorreport.st_deviation , comparatorreport.average_diff,weighting.ID,weighting.guid,weighting.ARCHIVAL,weighting.date,weighting.mass_in_g,weighting.tare_in_g,weighting.mass_in_unit,weighting.unit,weighting.UNIT_CAL,weighting.precision_act,weighting.precision_Cal,weighting.mass_is_stab,weighting.Mass_Air_Density_Correction,weighting.id_comparator,weighting.Cycle,weighting.AirDensity,weighting.THBTemperature,weighting.THBPressure,weighting.THBHumidity from weighting,comparatorreport where id_comparator=${correctArray[i]} AND weighting.id_comparator=comparatorreport.ID`, (error, results, fields)=>{
                    if (error) throw error;
                    let results_json = JSON.stringify(results);
                    let resultjson = JSON.parse(results_json);
                    workbook.xlsx.readFile('template.xlsx').then(()  => {
                    let worksheet = workbook.getWorksheet(1);
                    let environmental_conditions_start_column = 3;
                    let row_23 = 23;
                    let CountRow = 15;
                   
                    
                    let Conditiontemp = 55;           
                    for(let j=0; j<resultjson.length; j++)
                    {
                        let row_15 = worksheet.getRow(CountRow);
                        row_15.getCell(environmental_conditions_start_column).value = resultjson[j].THBTemperature;
                            console.log(CountRow);
                         
                        row_15.commit();
                        row_15 = row_23 + 7;
                        if(CountRow == Conditiontemp)
                        {
                            Conditiontemp = Conditiontemp + 55;
                            CountRow = CountRow + 15;
                        }
                        else 
                        {
                            CountRow = CountRow + 8;
                        }
                    }
                    workbook.xlsx.writeFile(`files/data.xlsx`);
                }).catch((err)=>{
                        console.log(err);
                })
            });
        }
    }
    }
const excel_obj = new excel();
excel_obj.initialize();