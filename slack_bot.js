var express = require('express');
var app = express();
var path    = require("path");
var fs = require('fs');
var GoogleSpreadsheet = require('google-spreadsheet');
var async = require('async');

//setInterval(() => 1, 1000 * 60 * 10);
app.set('port', (process.env.PORT || 5001));
app.set('view engine', 'html');
app.set('view options', {
  layout: false
});

app.use(express.static(__dirname + '/output'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

var http_server = app.listen(5001, function () {
  console.log('Server started on port %d', http_server.address().port);
});
//require('./route')(app, io);

// spreadsheet key is the long id in the sheets URL
var doc = new GoogleSpreadsheet('1PJQUM-HfIA7kj-fG-uIVkcFra3whG7Te2BOyPfzYrAs');
var sheet;
var row_ct = 0;

console.log ("start google spreadsheet");
function setAuth(step) {
  // see notes below for authentication instructions!
  var creds = require('./sirumobbot-d90905f5dd91.json');
  // OR, if you cannot save the file locally (like on heroku)
  var creds_json = {
   client_email: 'rapet87123@gmail.com',
   private_key: 'AIzaSyCsS3Nhxjk7DLREDlpqKELLCSL-YKO_U0I'
   };
  console.log ("get credentials");

  doc.useServiceAccountAuth(creds, step);
  console.log ("service auth is finished");
}

function getInfoAndWorksheets(step) {
  console.log ("start get info");
  doc.getInfo(function(err, info) {
    console.log('Loaded doc: '+info.title+' by '+info.author.email);
    sheet = info.worksheets[0];
    console.log('sheet 1: '+sheet.title+' '+sheet.rowCount+'x'+sheet.colCount);
    step();
  });
}

function workingWithRows(step) {
  console.log ("start workingWithRows");
  // google provides some query options
  //my_sheet.addRow( 2, { colname: 'col value'} );
  sheet.getRows({
    offset: 1,
    limit: 20,
    //orderby: 'col1'
  }, function( err, rows ){
    console.log('Read '+rows.length+' rows');
    //console.log (rows);

    // the row is an object with keys set by the column headers
    rows[0].colname = 'new val';
    rows[0]._cokwr = "Hey!!!";
    rows[0].save(); // this is async
    row_ct = rows.length;
    doc.addRow(sheet.id,{3:{4: { name: "a", val: 42 }, //'42' though tagged as "a"
      5: { name: "b", val: 21 }, //'21' though tagged as "b"
      6: "={{ a }}+{{ b }}"      //forumla adding row3,col4 with row3,col5 => '=D3+E3'
    }});

    // deleting a row
    //rows[0].del();  // this is async

    step();
  });

}

function workingWithCells(step) {
  console.log ("start working with cell");
  sheet.getCells({
    'start': row_ct,
    'min-row': row_ct+2,
    'max-row': row_ct+2,
    'return-empty': true,
    'limit' : 10,
    'max-col' : 4
  }, function(err, cells) {
    var cell = cells[0];
    console.log ("cell length : " + cells.length);
    console.log('Cell R'+cell.row+'C'+cell.col+' = '+cells.value);
    cell.value = 'Hey!!';
    /*cell.save(function(err){
      if (err)
        console.log (err);
    }); //async
    */
    /*
    // cells have a value, numericValue, and formula
    cell.value == '1'
    cell.numericValue == 1;
    cell.formula == '=ROW()';

    // updating `value` is "smart" and generally handles things for you
    cell.value = 123;
    cell.value = '=A1+B2'
    cell.save(); //async

    // bulk updates make it easy to update many cells at once
    cells[0].value = 1;
    cells[1].value = 2;
    cells[2].formula = '=A1+B1';
    //sheet.bulkUpdateCells(cells); //async
*/

    if (step)
      step(null);
      //step(null);
  });
}

function managingSheets(step) {
  console.log ("start working with manage sheet");
  doc.addWorksheet({
    title: 'my new sheet'
  }, function(err, sheet) {

    // change a sheet's title
    sheet.setTitle('new title'); //async

    //resize a sheet
    sheet.resize({rowCount: 50, colCount: 20}); //async

    sheet.setHeaderRow(['name', 'age', 'phone']); //async

    // removing a worksheet
    sheet.del(); //async

    step(null);
  });
}

function callbackhandler(err, results) {
  console.log('It came back with this ' + results);
}

async.series([
  setAuth, getInfoAndWorksheets, workingWithRows, workingWithCells//, managingSheets
], function(err, results){
  console.log ('err is ');
  console.log (err);
  console.log('Result of the whole run is ' + results);
});

function AppendNewAnswer(question, answer, text) {
  sheet.getCells({
    'start': row_ct,
    'min-row': row_ct+2,
    'max-row': row_ct+2,
    'return-empty': true,
    'limit' : 10,
    'max-col' : 4
  }, function(err, cells) {
    var cell = cells[0];
    cell.value = question;
    cell.save(function(err){
      if (err)
        console.log (err);
      else
      {
        var cell1 = cells[1];
        cell1.value = answer;
        cell1.save(function(err){
          if (err)
            console.log (err);
          else {
            var cell2 = cells[2];
            cell2.value = text;
            cell2.save(function(err) {
              if (err)
                console.log (err);
			  else 
				row_ct++;
            });
          }
        }); //async
      }
    }); //async
  });
}
