/**
 * 
 */
package org.app.myapp.testing;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.TextAutofit;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
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
	public static void main(String[] args) throws IOException {
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
//		  shape.getCTShape().getNvSpPr().getCNvPr().setId(shape.getCTShape().getNvSpPr().getCNvPr().getId()-1);

		   CTShapeProperties shapeProperties = shape.getCTShape().getSpPr();
		   CTLineProperties lineProperties = shapeProperties.getLn();

		   CTLineEndProperties lineEndProperties = org.openxmlformats.schemas.drawingml.x2006.main.CTLineEndProperties.Factory.newInstance();
		   lineEndProperties.setType(STLineEndType.TRIANGLE);
		   lineEndProperties.setLen(STLineEndLength.LG);
		   lineEndProperties.setW(STLineEndWidth.LG);

		   lineProperties.setTailEnd(lineEndProperties);
		   
		   
		   
		   
		   
		   final XSSFClientAnchor anc = new XSSFClientAnchor();
		

		   anc.setCol1(3);
		   anc.setDy1(1*XSSFShape.EMU_PER_PIXEL);
		   anc.setDx1(30*XSSFShape.EMU_PER_PIXEL);
		
//			   System.out.println("dx1=" +anc.getDx1() + " dy1="+anc.getDy1() +" dx2="+anc.getDx2() +" dy2="+anc.getDy2());
			
//			   float newWidthPixel = widthPixel * pourcentage;
//				anc.setDx1(Float.valueOf(newWidthPixel * XSSFShape.EMU_PER_PIXEL).intValue());
		 
		   anc.setCol2(4);
		   
		   anc.setRow1(3);
		   anc.setRow2(4);
		   
		   final XSSFSimpleShape simpleShape = ((XSSFDrawing) drawing).createSimpleShape(anc);
		   
		   simpleShape.getCTShape().getNvSpPr().getCNvPr().setName("ZONE DE TRAVAUX 1");
			simpleShape.setShapeType(ShapeTypes.RECT);
			simpleShape.setTextAutofit(TextAutofit.NORMAL);
			
			simpleShape.setText(new XSSFRichTextString("ZT1"));
//				simpleShape.setText("ZT1");
			simpleShape.setLineWidth(3);
			simpleShape.setFillColor(Color.CYAN.getRed(),
					Color.CYAN.getGreen(),
					Color.CYAN.getBlue());
			simpleShape.setTopInset(0);
			simpleShape.setLineStyleColor(Color.BLACK.getRed(), Color.BLACK.getGreen(), Color.BLACK.getBlue());
			simpleShape.setVerticalAlignment(VerticalAlignment.CENTER);
//				simpleShape.setTextAutofit(TextAutofit.NONE);
//				 simpleShape.getCTShape().getTxBody().getPArray(0).getPPr().setAlgn(STTextAlignType.CTR);
//				System.out.println(simpleShape.getCTShape().getTxBody().getPArray(0).getPPr());
			
		   

			 
//			double a = ((double)(((690 -686)*1000) + 0 - 131)/100) + 11 + (double)150/100;
//			double temp = ((double) 150) / 100;
//			int c = (300-131)/100;
//			int b = (int)a;
//			System.out.println(b);
			
//			Test1 t1 = new Test2();
			
			
try{
			   Date  dt = new Date();
			   SimpleDateFormat sdf = new SimpleDateFormat("MMddHHmmss");
			   String fname  = "workbook"+ (sdf.format(dt)).toString().replaceAll(":", "_")+".xlsx";
			   
			   FileOutputStream fileOut = new FileOutputStream(fname);
			   wb.write(fileOut);
			   System.out.println(fname + " généré");

			  } catch (IOException ioex) {
			  }
//
//	public int getNb1() {
//		return nb1;
//	}
//
//	public void setNb1(int nb1) {
//		this.nb1 = nb1;
//	}

}
}
// class Test21 extends Test1 {
//	private int nb2 ;
//
//	public int getNb2() {
//		return nb2;
//	}
//
//	public void setNb2(int nb2) {
//		this.nb2 = nb2;
//	}
//	
//}
//
//class Test3 extends Test1 {
//	private int nb3;
//
//	public int getNb3() {
//		return nb3;
//	}
//
//	public void setNb3(int nb3) {
//		this.nb3 = nb3;
//	}
//	
//}

