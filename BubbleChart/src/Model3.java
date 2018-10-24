import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.jfree.chart.ChartUtilities;
import org.jfree.data.general.DefaultPieDataset;

public class Model3 {

	public static void generateModel(File pptFileModel3) throws IOException {
		
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
		int leftMargin2 = 200;
		int topMargin1 = 50;
		int topMargin2 = pgy/2 + 20;
		int spacingLine = 15;
		int spacingParagraph = 30;
				
		//SET TITLE
		Utils.createSlideTitle(slide1, "Model 3");
		
		//INSERT SMILE IMAGE
	    File smileImage = new File("SmileImage.png"); 
	    Utils.createPresentationWithImage(smileImage, ppt, slide1, new Rectangle(leftMargin, topMargin1, (int)(44), (int)(40)));
	    
	    //INSERT TITLE
	  	XSLFTextBox shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin + 55, topMargin1, 160, 60));
	  	XSLFTextParagraph p = shape.addNewTextParagraph();	
	  	XSLFTextRun r1 = p.addNewTextRun();
	  	r1.setFontColor(new Color(225,125,125));
	  	r1.setText("Life - Liabilities AA");
	  	
	  	r1.setFontSize(18.);
	  	r1.setBold(true);
	  	
	  	 //INSERT TEXT
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + 50, 130, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Frequency");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	    //INSERT TEXT
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin2, topMargin1 + 50, 130, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Annual");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	  	
	    //INSERT VALUE ROBUSTNESS
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + 85, 130, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("3.6");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(12.);
	  	
	    //INSERT TEXT ROBUSTNESS
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + 100, 130, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Robustness");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	  	 //INSERT VALUE CONFORMITY
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + 135, 130, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("100%");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(12.);
	  	
	    //INSERT TEXT CONFORMITY
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + 150, 130, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Conformity");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	  	//INSERT STACKED BAR ROBUSTNESS
	  	File barStackedChartRobustness = new File("RobustnessBarStackedChart3.png");
	  	ChartUtilities.saveChartAsPNG(barStackedChartRobustness, Charts.stackedBarChartNoAxis(Datasets.createStackedBarDataset(3.6, 4.8, 4.0), "Robustness", "TYPE2"), 200, 60);
	  	Utils.createPresentationWithImage(barStackedChartRobustness, ppt, slide1, new Rectangle(leftMargin2 , topMargin1+85, 120, 20));
	  	
	    //INSERT STACKED BAR CONFORMITY
	  	File barStackedChartConformity = new File("ConformityBarStackedChart3.png");
	  	ChartUtilities.saveChartAsPNG(barStackedChartRobustness, Charts.stackedBarChartNoAxis(Datasets.createStackedBarDataset(99.81, 100, 95.00), "Conformity", "TYPE2"), 200, 60);
	  	Utils.createPresentationWithImage(barStackedChartRobustness, ppt, slide1, new Rectangle(leftMargin2 , topMargin1+135, 120, 20));
	  		
	  	int automaticTextDim = 175; 
	  	
	  	//INSERT TEXT ACTIVE CONTROLLS
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + automaticTextDim, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Active controls");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	    //INSERT VALUE ACTIVE CONTROLLS
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin2, topMargin1 + automaticTextDim, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("161");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(12.);
	  	
	    //INSERT TEXT QEngine
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin2+40, topMargin1 + automaticTextDim, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("(0 QEngine)");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	    //INSERT STACKED BAR ACTION CONTROLLS
	  	File activeControllsBarStacked = new File("ActiveControllsBarStackedChart.png");
	  	ChartUtilities.saveChartAsPNG(activeControllsBarStacked, Charts.stackedBarChartNoAxis(Datasets.createStackedBarDataset(90,71), "Action Controlls", "TYPE1"), 250, 60);
	  	Utils.createPresentationWithImage(activeControllsBarStacked, ppt, slide1, new Rectangle(leftMargin - 10, topMargin1+195, 220, 20));
	  	
	  	int automaticTextDim1 = 210;
	  	
	  	//INSERT AUTOMATIC TEXT
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + automaticTextDim1, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Automatic");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	  	
	  	//INSERT AUTOMATIC VALUE
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin+45, topMargin1 + automaticTextDim1, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("90");
	  	r1.setFontColor(Color.BLUE);
	  	r1.setFontSize(10.);
	  	
	  	
	    //INSERT AUTOMATIC TEXT
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin2+45, topMargin1 + automaticTextDim1, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Manual");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);  	
	  	
	  	//INSERT AUTOMATIC VALUE
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin2+80, topMargin1 + automaticTextDim1, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("71");
	  	r1.setFontColor(Color.MAGENTA);
	  	r1.setFontSize(10.);
	  	
		
	  	//GREEN BOX
	  	shape = slide1.createTextBox();
	  	shape.setFillColor(new Color(182,215,168));
	  	shape.setAnchor(new Rectangle(leftMargin, topMargin1 + 240, 80, 40));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Executed");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
		  	
	  	//RED BOX
	  	shape = slide1.createTextBox();
	  	shape.setFillColor(new Color(234,153,153));
	  	shape.setAnchor(new Rectangle(leftMargin+133, topMargin1 + 240, 80, 40));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("Not executed");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	    //INSERT CONFORMITY RING
	  	DefaultPieDataset dataset = new DefaultPieDataset( );
	    dataset.setValue( "OK", 75);  
	    dataset.setValue( "Below threshold", 6);
	    dataset.setValue( "Above treshold", 1);  
	  	File ringPlotConformity = new File("RingResult.png");
	  	ChartUtilities.saveChartAsPNG(ringPlotConformity, Charts.createRingChart(dataset, "", "TYPE2"), 200, 200);
	  	Utils.createPresentationWithImage(ringPlotConformity, ppt, slide1, new Rectangle(leftMargin -10, topMargin1 + 300, (int)(200*0.60), (int)(200*0.60)));	  	
	  	
	    //INSERT RESULTS TEXT
	  	shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin2+30, topMargin1 + 300, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
	  	r1.setText("82");
	  	r1.setBold(true);
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(16.);
	  	p.addLineBreak();
		r1 = p.addNewTextRun();
		r1.setText("Results Received");
	  	r1.setFontColor(Color.BLACK);
	  	r1.setFontSize(10.);
	  	
	  	//OK TEXT
		shape = slide1.createTextBox();
	  	shape.setAnchor(new Rectangle(leftMargin2+30, topMargin1 + 340, 120, 60));
	  	p = shape.addNewTextParagraph();	
	  	r1 = p.addNewTextRun();
		r1.setText("75 OK");
	  	r1.setFontColor(Color.GREEN);
	  	r1.setFontSize(10.);
		p.addLineBreak();
			  	
	  	//BELOW THRESHOLD
		r1 = p.addNewTextRun();
		r1.setText("6 Below threshold");
	  	r1.setFontColor(Color.YELLOW);
	  	r1.setFontSize(10.);
		p.addLineBreak();
		
	    //ABOVE THRESHOLD
		r1 = p.addNewTextRun();
		r1.setText("1 Above threshold");
	  	r1.setFontColor(Color.RED);
	  	r1.setFontSize(10.);
		p.addLineBreak();
	  	
	  	
		//CLOSE
		FileOutputStream out = new FileOutputStream(pptFileModel3);
		ppt.write(out);
		System.out.println("PowerPoint with image created!");

		out.close();
		
	}
}
