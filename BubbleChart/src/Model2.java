import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.util.IOUtils;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRelativeRect;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTable;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableCell;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTableProperties;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

public class Model2 {

	public static void generateModel(File pptFileModel2) throws IOException {

		    //CREATE SLIDE SHOW
		
			XMLSlideShow ppt = new XMLSlideShow();	
			XSLFSlideMaster slideMaster = ppt.getSlideMasters()[0];
			XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.BLANK);
			XSLFSlide slide1 = ppt.createSlide(slideLayout);
			
			//SET TITLE
			Utils.createSlideTitle(slide1, "Model 2");
			
			int numColumns = 3;
			int numRows = 3;
					
			XSLFTable tbl = slide1.createTable();
					
			Dimension pgsize = ppt.getPageSize();
		    int pgx = pgsize.width; //slide width in points
		    int pgy = pgsize.height; //slide height in points
			
			tbl.setAnchor(new Rectangle(25, 50, pgx/2 - 100, 25*(numRows+1)));
						
			XSLFTableRow headerRow = tbl.addRow();
			headerRow.setHeight(15);
			String[] headers = {"Area","Robustness","Conformity"};
			
			// header
			for (int i = 0; i < numColumns; i++) {
				XSLFTableCell th = headerRow.addCell();
				th.setBorderRight(2);
				th.setBorderRightColor(new Color(255,255,255));
				XSLFTextParagraph p = th.addNewTextParagraph();
				p.setTextAlign(TextAlign.CENTER);
				XSLFTextRun r = p.addNewTextRun();
				r.setText(headers[i]);
				r.setBold(true);
				r.setFontSize(10);
				th.setFillColor(new Color(79, 129, 189));			
				// th.setBorderWidth(BorderEdge.bottom, 2.0);
				// th.setBorderColor(BorderEdge.bottom, Color.white);

				tbl.setColumnWidth(i, 150); // all columns are equally sized
			}

			//rows        
			for (int rownum = 0; rownum < numRows; rownum++) {
				XSLFTableRow tr = tbl.addRow();
				tr.setHeight(15);
							
				XSLFTableCell cell = tr.addCell();
				XSLFTextParagraph p = cell.addNewTextParagraph();
				XSLFTextRun r = p.addNewTextRun();
				r.setText("90");
				r.setFontSize(10);
				r.setBold(true);
				cell.setFillColor(new Color(251,248,185));
				cell.setBorderBottom(1.5);   //for width of border
				cell.setBorderBottomColor(new Color(255,255,255));
				cell.setBorderRight(2);
				cell.setBorderRightColor(new Color(255,255,255));
				 
				r = p.addNewTextRun();
				r.setText(" CONTROLS");
				r.setFontSize(10);
				p.addLineBreak();
				
				
				if(rownum == 0){
				
				p.addLineBreak();
				r = p.addNewTextRun();
				r.setText("Robustness");
				r.setFontSize(8);
				p.addLineBreak();
				
				r = p.addNewTextRun();			
				r.setText("1.3");
				r.setFontSize(10);
				r.setBold(true);
				p.addLineBreak();
				p.addLineBreak();
				
				r = p.addNewTextRun();			
				r.setText("Conformity");
				r.setFontSize(8);
				p.addLineBreak();
				
				r = p.addNewTextRun();			
				r.setText("100%");
				r.setFontSize(10);
				r.setBold(true);
				p.addLineBreak();
				p.addLineBreak();
				
				}
							
				//2nd column
				cell = tr.addCell();
				cell.setFillColor(new Color(251,248,185));
				cell.setBorderBottom(1);   //for width of border
				cell.setBorderBottomColor(new Color(255,255,255));
				cell.setBorderRight(2);
				cell.setBorderRightColor(new Color(255,255,255));
				p = cell.addNewTextParagraph();
				r = p.addNewTextRun();
				r.setFontSize(10);
				r.setText("99");
				
				//3rd column
				cell = tr.addCell();
				cell.setFillColor(new Color(251,248,185));
				cell.setBorderBottom(1);   //for width of border
				cell.setBorderBottomColor(new Color(255,255,255));
				cell.setBorderRight(2);
				cell.setBorderRightColor(new Color(255,255,255));
				p = cell.addNewTextParagraph();
				r = p.addNewTextRun();
				r.setFontSize(10);
				r.setText("78.2");
				
				//TEST IMAGE IN TABLE
				//  ????
				
			}
			
			//CLOSE
			FileOutputStream out = new FileOutputStream(pptFileModel2);
			ppt.write(out);
			System.out.println("PowerPoint with image created!");

			out.close();
			
		}
	
}
