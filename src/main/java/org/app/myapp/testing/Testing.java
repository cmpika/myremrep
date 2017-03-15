/**
 * 
 */
package org.app.myapp.testing;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

			   lineProperties.setTailEnd(lineEndProperties);

			   FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
			   wb.write(fileOut);

			  } catch (IOException ioex) {
			  }

	}

}
