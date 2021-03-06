const express = require('express');
const app = express();
const mysql = require('mysql');
var Excel = require('exceljs');
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
    cyclesiterable(correctArray,cycles)
    {
        const connection = mysql.createConnection(this.configure);
        var workbook =  new Excel.Workbook();
        workbook.xlsx.readFile('template.xlsx').then(()  => {
            const worksheet = workbook.getWorksheet(1);
            let CellPropertyLeft = 3;
            let RowCompareStart = 15;
            let CellCompare = 4;
            let RowTemperatureStart = 15;
        for(let i=0; i<correctArray.length; i++)
        {
            if(i>0)
            {
                RowCompareStart =RowCompareStart + 55;
                RowTemperatureStart = RowTemperatureStart + 55;
            }
            for(let j=1; j<=6; j++) {
            connection.query(`SELECT comparatorreport.st_deviation , comparatorreport.average_diff,weighting.ID,weighting.guid,weighting.ARCHIVAL,weighting.date,weighting.mass_in_g,weighting.tare_in_g,weighting.mass_in_unit,weighting.unit,weighting.UNIT_CAL,weighting.precision_act,weighting.precision_Cal,weighting.mass_is_stab,weighting.Mass_Air_Density_Correction,weighting.id_comparator,weighting.Cycle,weighting.AirDensity,weighting.THBTemperature,weighting.THBPressure,weighting.THBHumidity from weighting,comparatorreport where id_comparator=${correctArray[i]} AND weighting.id_comparator=comparatorreport.ID AND weighting.Cycle=${j}`, (error, results, fields)=>{
                if (error) throw error;
                let results_json = JSON.stringify(results);
                let resultjson = JSON.parse(results_json);
                for(let l=0; l<=3; l++)
                {
                    let RowCompare
                    let RowPropertyLeft
                    RowCompare = worksheet.getRow(RowCompareStart);
                    RowCompare.getCell(CellCompare).value = resultjson[l].mass_in_g;
                    RowCompare.commit(); 
                    CellCompare = CellCompare+1;
                    if(CellCompare>7)
                    {
                        CellCompare = 4;
                        RowCompareStart = RowCompareStart+7;
                    }
                    RowPropertyLeft = worksheet.getRow(RowTemperatureStart);
                    RowPropertyLeft.getCell(CellPropertyLeft).value = resultjson[l].THBTemperature;
                    RowPropertyLeft.commit();
                    if(j>0)
                    {
                    RowTemperatureStart = RowTemperatureStart+8;
                    }
                }
        });
        }
            }
            }).catch((err)=>{
                console.log(err);
            })
            workbook.xlsx.writeFile(`files/data.xlsx`);
    }
    }
const excel_obj = new excel();
excel_obj.initialize();