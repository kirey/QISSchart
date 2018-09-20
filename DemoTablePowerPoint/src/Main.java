import java.awt.Color;
import java.awt.Rectangle;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;


public class Main {

	public static void main(String[] args) throws IOException {
		try {
			XMLSlideShow ppt = new XMLSlideShow();
			XSLFSlide slide = ppt.createSlide();

			XSLFTable tbl = slide.createTable();
			tbl.setAnchor(new Rectangle(50, 50, 450, 300));

			int numColumns = 3;
			int numRows = 5;
			XSLFTableRow headerRow = tbl.addRow();
			headerRow.setHeight(50);
			// header
			for (int i = 0; i < numColumns; i++) {
				XSLFTableCell th = headerRow.addCell();
				XSLFTextParagraph p = th.addNewTextParagraph();
				p.setTextAlign(TextAlign.CENTER);
				XSLFTextRun r = p.addNewTextRun();
				r.setText("Header " + (i + 1));
				r.setBold(true);
				r.setFontColor(Color.white);
				th.setFillColor(new Color(79, 129, 189));
//				th.setBorderWidth(BorderEdge.bottom, 2.0);
//				th.setBorderColor(BorderEdge.bottom, Color.white);

				tbl.setColumnWidth(i, 150); // all columns are equally sized
			}

			// rows

			for (int rownum = 0; rownum < numRows; rownum++) {
				XSLFTableRow tr = tbl.addRow();
				tr.setHeight(50);
				// header
				for (int i = 0; i < numColumns; i++) {
					XSLFTableCell cell = tr.addCell();
					XSLFTextParagraph p = cell.addNewTextParagraph();
					XSLFTextRun r = p.addNewTextRun();

					r.setText("Cell " + (i + 1));
					if (rownum % 2 == 0)
						cell.setFillColor(new Color(208, 216, 232));
					else
						cell.setFillColor(new Color(233, 247, 244));

				}
			}

			try (FileOutputStream out = new FileOutputStream("table.pptx")) {
				ppt.write(out);
			}
		}catch(Exception ex) {
			ex.printStackTrace();
		}
	}
}
