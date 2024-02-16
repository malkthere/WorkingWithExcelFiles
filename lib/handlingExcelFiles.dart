import 'dart:io';
import 'package:file_picker/file_picker.dart';

import 'package:flutter/material.dart';
import 'package:flutter_excel/excel.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:path/path.dart';
import 'package:flutter/foundation.dart';


class Excelfiles {
  int _counter = 0;

  //get selectedExcel => null;
  var excel = Excel.createExcel();

  void createexcelfile() async{
    Sheet sheetObject = excel['فاتورة 01'];
    var cell = sheetObject.cell(CellIndex.indexByString("A1"));
    cell.value = 8; // dynamic values support provided;
    List<String> data = ["Mr","’mazin", "Alkathiri"];
    // sheetObject = selectedExcel["Sheet1"];
    sheetObject.appendRow(data);
    saveExcel();

  }
  saveExcel() async {
    //request for storage permission
    var res = await Permission.storage.request();
    var fileBytes = excel.save();
    //var directory = await getApplicationDocumentsDirectory();

    File(join("/storage/emulated/0/Download/billNo03.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes!);
    //"/storage/emulated/0/Download/"  download folder address
    //excel2.xlsx is the file name "feel free to change the file name to anything you want"



  }

  void readexcelfile(File? file) async {
    //var file = "/storage/emulated/0/Download/billNo02.xlsx";
    var bytes = file!.readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    for (var table in excel.tables.keys) {
      // print(table); //sheet Name
      ////  print(excel.tables[table]?.maxCols);
      // print(excel.tables[table]?.maxRows);
      int counter=0;
      for (var row in excel.tables[table]!.rows) {
        // print("$row");
        if(counter==1){
          int counter2=0;
          for (var cell in row) {
            if(counter2==4) {
              //  print('cell ${cell!.rowIndex}/${cell.columnIndex}');
              if (cell != null)
                print(cell!.value);
            }
            counter2++;
          }
        }
        if(counter==2){
          int counter2=0;
          for (var cell in row) {
            if(counter2==4) {
              //  print('cell ${cell!.rowIndex}/${cell.columnIndex}');
              if (cell != null)
                print(cell!.value);
            }
            counter2++;
          }
        }
        if(counter>6){
          int counter2=0;
          for (var cell in row) {
            if(counter2==1) {
              //  print('cell ${cell!.rowIndex}/${cell.columnIndex}');
              if (cell != null)
                print(cell!.value);
            }
            if(counter2==2) {
              //  print('cell ${cell!.rowIndex}/${cell.columnIndex}');
              if (cell != null)
                print(cell!.value);
            }
            counter2++;
          }
        }
        counter++;
      }
    }
  }

}