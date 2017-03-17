/**
 * 
 */
package org.app.myapp.testing;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.TextAutofit;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineEndProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndLength;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndType;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndWidth;

/**
 * @author cmpika
 *
 */
public class Testing {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		 try {

			   Workbook wb = new XSSFWorkbook();

			   Sheet sheet = wb.createSheet("Sheet1");

			   CreationHelper helper = wb.getCreationHelper();
			   Drawing drawing = sheet.createDrawingPatriarch();
			   
			   
			   
			   
			   
			   final XSSFClientAnchor anc = new XSSFClientAnchor();
			   anc.setCol1(3);
			   anc.setDx1(1);
				
			   anc.setDx2(6);
			   anc.setCol2(7);
			   
			   anc.setRow1(1);
			   anc.setRow2(1 + 2);
			   
			   final XSSFSimpleShape simpleShape = ((XSSFDrawing) drawing).createSimpleShape(anc);
			   
			   simpleShape.getCTShape().getNvSpPr().getCNvPr().setName("PJEL_DEMANDE");
				simpleShape.setShapeType(ShapeTypes.RECT);
				simpleShape.setTextAutofit(TextAutofit.NORMAL);
				simpleShape.setText("toto");
				simpleShape.setLineWidth(3);
				simpleShape.setFillColor(Color.YELLOW.getRed(),
						Color.YELLOW.getGreen(),
						Color.YELLOW.getBlue());
				simpleShape.setTopInset(0);
				simpleShape.setLineStyleColor(Color.BLACK.getRed(), Color.BLACK.getGreen(), Color.BLACK.getBlue());
				
			   
			   
			   

			   ClientAnchor anchor = helper.createClientAnchor();
			   anchor.setCol1(2);
			   anchor.setRow1(2); 
			   anchor.setCol2(5);
			   anchor.setRow2(5); 

			   XSSFSimpleShape shape = ((XSSFDrawing)drawing).createSimpleShape((XSSFClientAnchor)anchor);
			   shape.setShapeType(ShapeTypes.LINE);
			   shape.setLineWidth(1.5);
			   shape.setLineStyle(3);
			   shape.setLineStyleColor(0,0,255);

			//apache POI sets first shape Id to 1. It should be 0.
			shape.getCTShape().getNvSpPr().getCNvPr().setId(shape.getCTShape().getNvSpPr().getCNvPr().getId()-1);

			   CTShapeProperties shapeProperties = shape.getCTShape().getSpPr();
			   CTLineProperties lineProperties = shapeProperties.getLn();

			   CTLineEndProperties lineEndProperties = org.openxmlformats.schemas.drawingml.x2006.main.CTLineEndProperties.Factory.newInstance();
			   lineEndProperties.setType(STLineEndType.TRIANGLE);
			   lineEndProperties.setLen(STLineEndLength.LG);
			   lineEndProperties.setW(STLineEndWidth.LG);

			   lineProperties.setHeadEnd(lineEndProperties);

			   FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
			   wb.write(fileOut);
			   System.out.println("workbook.xlsx généré");

			  } catch (IOException ioex) {
			  }

	}

}
