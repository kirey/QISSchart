import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.jfree.chart.ChartUtilities;

public class Model1 {
	
	public static void generateModel1(File pptFileModel1) throws IOException {
		
		//CREATE SLIDE SHOW
		
		XMLSlideShow ppt = new XMLSlideShow();	
		XSLFSlideMaster slideMaster = ppt.getSlideMasters()[0];
		XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.BLANK);
		XSLFSlide slide1 = ppt.createSlide(slideLayout);
		XSLFSlide slide2 = ppt.createSlide();
		
		Dimension pgsize = ppt.getPageSize();
	    int pgx = pgsize.width; //slide width in points
	    int pgy = pgsize.height; //slide height in points
	    int leftMargin = 100;
	    int topMargin1 = 50;
	    int topMargin2 = pgy/2 + 20;
		
		//SET TITLE
		Utils.createSlideTitle(slide1, "Model 1");
	
		//CREATE AND ADD BUBBLE CHART
		File bubbleChart = new File("BubbleChart.png");
		int width = 560; /* Width of the image */
		int height = 370; /* Height of the image */
		ChartUtilities.saveChartAsPNG(bubbleChart, Charts.createBubbleChart(Datasets.createDataset(), "", "Robustness", "Conformity"), width, height);
	    Utils.createPresentationWithImage(bubbleChart, ppt, slide1, new Rectangle(leftMargin, topMargin1, (int)(560*0.60), (int)(370*0.60)));
		
		//INSERT SMILE IMAGE
	    File smileImage = new File("SmileImage.png"); 
	    Utils.createPresentationWithImage(smileImage, ppt, slide1, new Rectangle(leftMargin, topMargin2 + 60 - 20, (int)(44), (int)(40)));
		
	    
		//INSERT RING ROBUSTNESS
		File ringPlotRobustness = new File("RingPlotRobustness.png");
		ChartUtilities.saveChartAsPNG(ringPlotRobustness, Charts.createRingChart(Datasets.createDatasetPie(4.8, 3.6), "Robustness", "TYPE1"), 200, 200);
		Utils.createPresentationWithImage(ringPlotRobustness, ppt, slide1, new Rectangle(leftMargin + 50, topMargin2, (int)(200*0.60), (int)(200*0.60)));
		
		//INSERT TEXT BELOW RING ROBUSTNESS
		XSLFTextBox shape = slide1.createTextBox();
		shape.setAnchor(new Rectangle(leftMargin + 55, pgy/2 + 130, 120, 60));
		XSLFTextParagraph p = shape.addNewTextParagraph();	
		XSLFTextRun r1 = p.addNewTextRun();
		r1.setText("Robustness vs target (default)");
		r1.setFontColor(Color.BLACK);
		r1.setFontSize(8.);
		
		//INSERT STACKED BAR ROBUSTNESS
		File barStackedChartRobustness = new File("RobustnessBarStackedChart.png");
		ChartUtilities.saveChartAsPNG(barStackedChartRobustness, Charts.createStackedBarChart(Datasets.createStackedBarDataset(3.6, 4.8, 4.0), "Robustness"), 200, 70);
		Utils.createPresentationWithImage(barStackedChartRobustness, ppt, slide1, new Rectangle(leftMargin + 50, topMargin2+165, 120, 42));
		
		
		
		//INSERT CONFORMITY RING
		File ringPlotConformity = new File("RingPlotConformity.png");
		ChartUtilities.saveChartAsPNG(ringPlotConformity, Charts.createRingChart(Datasets.createDatasetPie(100.0, 99.94), "Conformity", "TYPE1"), 200, 200);
		Utils.createPresentationWithImage(ringPlotConformity, ppt, slide1, new Rectangle(pgx/2, topMargin2, (int)(200*0.60), (int)(200*0.60)));

		//INSERT TEXT BELOW RING CONFORMITY
		shape = slide1.createTextBox();
		shape.setAnchor(new Rectangle(pgx/2 + 5, topMargin2 + 130, 120, 60));
		p = shape.addNewTextParagraph();	
		r1 = p.addNewTextRun();
		r1.setText("Conformity vs target (default)");
		r1.setFontColor(Color.BLACK);
		r1.setFontSize(8.);
		
		//INSERT STACKED BAR CONFORMITY
		File barStackedChartConformity = new File("ConformityBarStackedChart.png");
		ChartUtilities.saveChartAsPNG(barStackedChartRobustness, Charts.createStackedBarChart(Datasets.createStackedBarDataset(3.6, 4.8, 4.0), "Conformity"), 200, 70);
		Utils.createPresentationWithImage(barStackedChartRobustness, ppt, slide1, new Rectangle(pgx/2, topMargin2+165, 120, 42));
		
		
		//INSERT TABLE ON SECOND SLIDE
		createTableOnSlide(ppt, slide2);
		
		
		//CLOSE
		FileOutputStream out = new FileOutputStream(pptFileModel1);
		ppt.write(out);
		System.out.println("PowerPoint with image created!");

		out.close();

		
		
	}
	
	private static void createTableOnSlide(XMLSlideShow ppt, XSLFSlide slide) throws IOException {

		// Add table to ppt on the same page
		
		int numColumns = 3;
		int numRows = demoData().size();
			
		XSLFTable tbl = slide.createTable();
		
		Dimension pgsize = ppt.getPageSize();
	    int pgx = pgsize.width; //slide width in points
	    int pgy = pgsize.height; //slide height in points
		
		tbl.setAnchor(new Rectangle(50, 50, pgx/2, 25*(numRows+1)));
		
		XSLFTableRow headerRow = tbl.addRow();
		headerRow.setHeight(15);
		String[] headers = {"Area","Robustness","Conformity"};
		
		// header
		for (int i = 0; i < numColumns; i++) {
			XSLFTableCell th = headerRow.addCell();
			XSLFTextParagraph p = th.addNewTextParagraph();
			p.setTextAlign(TextAlign.CENTER);
			XSLFTextRun r = p.addNewTextRun();
			r.setText(headers[i]);
			r.setBold(true);
			r.setFontSize(10);
			th.setFillColor(new Color(248, 248, 248));
			// th.setBorderWidth(BorderEdge.bottom, 2.0);
			// th.setBorderColor(BorderEdge.bottom, Color.white);

			tbl.setColumnWidth(i, 150); // all columns are equally sized
		}

		//rows        
		for (int rownum = 0; rownum < numRows; rownum++) {
			XSLFTableRow tr = tbl.addRow();
			tr.setHeight(15);
		
			DemoDataObj dataRow = demoData().get(rownum);
			
			XSLFTableCell cell = tr.addCell();
			XSLFTextParagraph p = cell.addNewTextParagraph();
			XSLFTextRun r = p.addNewTextRun();
			r.setText(dataRow.getCompany());
			r.setFontSize(10);
			p.addLineBreak();
			
			cell = tr.addCell();
			p = cell.addNewTextParagraph();
			r = p.addNewTextRun();
			r.setFontSize(10);
			r.setText(dataRow.getRobustness());
			
			cell = tr.addCell();
			p = cell.addNewTextParagraph();
			r = p.addNewTextRun();
			r.setFontSize(10);
			r.setText(dataRow.getConformity());
			
		}
		
	}
	
	public static List<DemoDataObj> demoData(){
		
		List<DemoDataObj> demoData = new ArrayList<DemoDataObj>();
		
		DemoDataObj data = new DemoDataObj();
		data.setCompany("AA");
		data.setConformity("99.81");
		data.setRobustness("4.1");
		demoData.add(data);
		
		data = new DemoDataObj();
		data.setCompany("GI");
		data.setConformity("100");
		data.setRobustness("2.7");
		demoData.add(data);
		
		data = new DemoDataObj();
		data.setCompany("GT");
		data.setConformity("100");
		data.setRobustness("4.1");
		demoData.add(data);
		
		return demoData;
		
	}

}
