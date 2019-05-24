const express = require('express');
const app = express();
const mysql = require('mysql');
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
        connection.connect();
            connection.query('SELECT id_comparator from weighting ORDER BY id_comparator DESC LIMIT 10', (error, results, fields) => {
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
                    connection.end();
                    that.cyclesiterable(correctArray);
                  })
        });
    }
    cyclesiterable(correctArray)
    {
        const connection = mysql.createConnection(this.configure);
        connection.connect();
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
                    for(let j=0; i<resultjson.length; i++)
                    {
                        row_23.getCell(environmental_conditions_start_column).value = resultjson[j].THBTemperature;
                        row_23.commit();
                        row_23 = row_23 + 7;
                    }
                    
                    workbook.xlsx.writeFile(`files/data.xlsx`);
                    connection.end();
                }).catch((err)=>{
                        console.log(err);
                })
            });
        }
    }
    }
const excel_obj = new excel();
excel_obj.initialize();