/**
 * 
 */
package org.app.myapp.testing;

/**
 * @author cmpika
 *
 */
	
	import java.awt.Color;
	import java.io.*;

import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;
	import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextAlignType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextAlignType.Enum; 

	public class Test2 {

	    static String fileName = "TestWorkbook.xlsx";

	    public static void main(String[] args) throws Exception {
	        XSSFWorkbook wb = new XSSFWorkbook();
	        XSSFSheet sht = wb.createSheet();
	        File file = new File(fileName);
//	        int colStart = 1;
	        int colStart = 5;

	        XSSFDrawing draw = sht.createDrawingPatriarch();

	        XSSFShapeGroup group = draw.createGroup(draw.createAnchor(0, 0, 0, 0, colStart, 11, colStart + 6, 11+7));
	        group.setCoordinates(colStart, 11, colStart + 6, 11+7);

	        XSSFTextBox tb1 = group.createTextbox(new XSSFChildAnchor(5, 0, 9, 10));
	        tb1.setLineStyleColor(0, 0, 0);
	        tb1.setLineWidth(2);
	        Color col = Color.cyan;
	        tb1.setFillColor(col.getRed(), col.getGreen(), col.getBlue());

	        XSSFRichTextString address = new XSSFRichTextString("TextBox string 1\nHas three\nLines to it");
	        tb1.setText(address);    
	        tb1.setVerticalAlignment(VerticalAlignment.CENTER);
	        
	        CTTextCharacterProperties rpr = tb1.getCTShape().getTxBody().getPArray(0).getRArray(0).getRPr();
	        rpr.addNewLatin().setTypeface("Trebuchet MS");
	        rpr.setSz(900); // 9 pt
	        col = Color.black;
	        rpr.addNewSolidFill().addNewSrgbClr().setVal(new byte[]{(byte)col.getRed(),(byte)col.getGreen(),(byte)col.getBlue()});

	        FileOutputStream fout = new FileOutputStream(file);
	        wb.write(fout);
	        fout.close();
	    }
	}

