import java.awt.Color;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

public class Utils {

	public static void createSlideTitle(XSLFSlide slide, String text) {
		//ADD TEXT FOR TITLE
		XSLFTextBox shape = slide.createTextBox();
		shape.setAnchor(new Rectangle(250, 5, 150, 25));
		XSLFTextParagraph p = shape.addNewTextParagraph();
		p.setTextAlign(TextAlign.CENTER);

		XSLFTextRun r1 = p.addNewTextRun();
		r1.setText(text);
		r1.setFontColor(Color.BLACK);
		r1.setFontSize(24.);
		
	}
	
	public static void createPresentationWithImage(File image, XMLSlideShow ppt, XSLFSlide slide, Rectangle rectangle)
			throws IOException {
		
		byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
		int idx = ppt.addPicture(picture, XSLFPictureData.PICTURE_TYPE_PNG);
		XSLFPictureShape pic = slide.createPicture(idx);
		pic.setAnchor(rectangle);
		
	}
}
