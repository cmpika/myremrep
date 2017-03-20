/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.app.myapp.testing;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

/**
 * Demonstrates how to create a simple table using Apache POI.
 */
public class CreateTable {
        
    @SuppressWarnings("deprecation")
	public static void main(String[] args) throws IOException {
        
//    	Workbook wb = new XSSFWorkbook();
//        XSSFSheet sheet = (XSSFSheet) wb.createSheet();
//        
//        //Create 
//        XSSFTable table = sheet.createTable();
//        table.setDisplayName("Test");       
//        CTTable cttable = table.getCTTable();
        
        //Style configurations
//        CTTableStyleInfo style = cttable.addNewTableStyleInfo();
        //style.setName("TableStyleMedium19");
        //style.setName("TableStyleMedium25");
//      style.setName("TableStyleMedium22");
//      style.setName("TableStyleMedium4");
//        style.setName("TableStyleMedium15");
      
//        style.setShowColumnStripes(false);
//        style.setShowRowStripes(true);
        
        //Set which area the table should be placed in
//        AreaReference reference = new AreaReference(new CellReference(1, 1), 
//                new CellReference(13,4));
//        cttable.setRef(reference.formatAsString());
//        cttable.setId(1);
//        cttable.setName("Test");
//        
//        CTTableColumns columns = cttable.addNewTableColumns();
//        columns.setCount(3);
//        CTTableColumn column;
//        XSSFRow row;
//        XSSFCell cell;
//        for(int i=1; i<5; i++) {
//            //Create column
//            column = columns.addNewTableColumn();
//            column.setId(i+1);
//            //Create row
//            row = sheet.createRow(i);
//            for(int j=1; j<5; j++) {
//                //Create cell
//                cell = row.createCell(j);
//                    cell.setCellValue("test");
//            }
//        }
//        
        
        
        
//        sheet.addMergedRegion(new CellRangeAddress(
//	            1, //first row (0-based)
//	            1, //last row  (0-based)
//	            1, //first column (0-based)
//	            2  //last column  (0-based))
//	            ));
        
        
//        
//        
//        FileOutputStream fileOut = new FileOutputStream("ooxml-table.xlsx");
//        wb.write(fileOut);
//        fileOut.close();
//        wb.close();
        
        File f = new File("test.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("no blue");

        // set the color of the cell
        XSSFCellStyle style = wb.createCellStyle();
        XSSFColor myColor = new XSSFColor(Color.WHITE);
        style.setFillForegroundColor(myColor);
        style.setFillBackgroundColor(myColor);
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderTop(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        cell.setCellStyle(style); // this command seems to fail

        try {
            FileOutputStream fos = new FileOutputStream(f);
            wb.write(fos);
            wb.close();
            fos.flush();
            fos.close();
        }catch(Exception e){
            e.printStackTrace();
        }

    }
}