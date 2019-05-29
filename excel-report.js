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
    cyclesiterable(correctArray,cycles)
    {
        const connection = mysql.createConnection(this.configure);
        for(let i=0; i<correctArray.length; i++){
            for(let j=1; j<=6; j++) {
            connection.query(`SELECT comparatorreport.st_deviation , comparatorreport.average_diff,weighting.ID,weighting.guid,weighting.ARCHIVAL,weighting.date,weighting.mass_in_g,weighting.tare_in_g,weighting.mass_in_unit,weighting.unit,weighting.UNIT_CAL,weighting.precision_act,weighting.precision_Cal,weighting.mass_is_stab,weighting.Mass_Air_Density_Correction,weighting.id_comparator,weighting.Cycle,weighting.AirDensity,weighting.THBTemperature,weighting.THBPressure,weighting.THBHumidity from weighting,comparatorreport where id_comparator=${correctArray[i]} AND weighting.id_comparator=comparatorreport.ID AND weighting.Cycle=${j}`, (error, results, fields)=>{
                if (error) throw error;
                let results_json = JSON.stringify(results);
                let resultjson = JSON.parse(results_json);
                for(let l=0; l<=3; l++)
                {
                    
                }
        });
    
        }
    }
    }
    }
const excel_obj = new excel();
excel_obj.initialize();