import java.awt.Color;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.AxisLocation;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.XYPlot;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;
import org.jfree.data.xy.DefaultXYZDataset;
import org.jfree.data.xy.XYZDataset;

public class Main {

	/**
	 * @param args
	 * @throws IOException
	 */
	final static Color STANDBY_COLOR = Color.DARK_GRAY;
    final static Color HEATING_COLOR = Color.ORANGE;
    final static Color HOLDING_COLOR = Color.YELLOW;
    final static Color COOLING_COLOR = Color.CYAN;
    final static Color LOWERING_COLOR = Color.GREEN;
        
	public static void main(String[] args) throws IOException {
		
		
		final List<DemoDataObj> demoData = demoData();
		
		//CREATE SLIDE SHOW
		XMLSlideShow ppt = new XMLSlideShow();	
		XSLFSlideMaster slideMaster = ppt.getSlideMasters()[0];
		XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.BLANK);
		XSLFSlide slide = ppt.createSlide(slideLayout);
		//XSLFSlide slide1 = ppt.createSlide(slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT));
		XSLFSlide slide2 = ppt.createSlide();
		
		
		//CREATING CHART PNGS
		File bubbleChart = new File("BubbleChart.png");
		File pieChartConformity = new File("ConformityPieChart.png");		
		File pieChartRobustness = new File("RobustnessPieChart.png");	
		File barChartRobustness = new File("RobustnessBarChart.png");
		File barStackedChartRobustness = new File("RobustnessBarStackedChart.png");
		File barStackedChartConformity = new File("ConformityBarStackedChart.png");
		File ringPlot = new File("RingPlot.png");

		Dimension pgsize = ppt.getPageSize();
	    int pgx = pgsize.width; //slide width in points
	    int pgy = pgsize.height; //slide height in points
		
		int width = 560; /* Width of the image */
		int height = 370; /* Height of the image */
		ChartUtilities.saveChartAsPNG(bubbleChart, createBubbleChart(), width, height);
		ChartUtilities.saveChartAsPNG(pieChartRobustness, createPieChart(createDatasetPie(4.8, 3.6), "Robustness"), 330, 330);
		ChartUtilities.saveChartAsPNG(pieChartConformity, createPieChart(createDatasetPie(100.0, 99.94), "Conformity"), 330, 330);
		ChartUtilities.saveChartAsPNG(barChartRobustness, createBarChart(createProgressBarDataset(), "Robustness"), 200, 350);
		ChartUtilities.saveChartAsPNG(barStackedChartRobustness, createStackedBarChart(createStackedBarDataset(3.6, 4.8, 4.0), "Robustness"), 350, 150);
		ChartUtilities.saveChartAsPNG(barStackedChartConformity, createStackedBarChart(createStackedBarDataset(99.94, 100, 95.0), "Conformity"), 350, 150);
		ChartUtilities.saveChartAsPNG(ringPlot, createRingChart(createDatasetPie(100.0, 65.0), "Ring Test"), 200, 200);
		//createStackedBarDataset
		System.out.println("Image rendered!");

		
		//ADDING ELEMENTS ON SLIDE
		File file = new File("PowerPointBubble.pptx");

		createSlideTitle(slide, "Model 1");
		createPresentationWithImage(bubbleChart, file, ppt, slide, new Rectangle(pgx/2, 35, (int)(560*0.60), (int)(370*0.60)));
		createPresentationWithImage(pieChartRobustness, file, ppt, slide, new Rectangle(35, pgy/2-35, (int)(330*0.6), (int)(330*0.6)));
		createPresentationWithImage(pieChartConformity, file, ppt, slide, new Rectangle(pgx/2, pgy/2-35, (int)(330*0.6), (int)(330*0.6)));
		//createPresentationWithImage(barChartRobustness, file, ppt, slide, new Rectangle(35, pgy/2, (int)(200*0.60), (int)(350*0.60)));
		createPresentationWithImage(barStackedChartRobustness, file, ppt, slide, new Rectangle(35, pgy/2+165, (int)(350*0.60), (int)(150*0.60)));
		createPresentationWithImage(barStackedChartConformity, file, ppt, slide, new Rectangle(pgx/2, pgy/2+165, (int)(350*0.60), (int)(150*0.60)));
		createTableOnSlide(file, ppt, slide);
		createPresentationWithImage(ringPlot, file, ppt, slide2, new Rectangle(35, pgy/2, (int)(200*0.60), (int)(200*0.60)));
		Model2.createTableModel2(file, ppt, slide2);

		//////TEST///// inserting image to layout content
		/*
		//String img = "BubbleChart.png";
		//BufferedImage myPicture = ImageIO.read(new File(imagePath));
		
		//converting it into a byte array
		byte[] pictureData = IOUtils.toByteArray(new FileInputStream(bubbleChart));
		byte[] pictureData2 = IOUtils.toByteArray(new FileInputStream(pieChartConformity));

		//adding the image to the presentation
		int idx = ppt.addPicture(pictureData, XSLFPictureData.PICTURE_TYPE_PNG);
		int idx2 = ppt.addPicture(pictureData2, XSLFPictureData.PICTURE_TYPE_PNG);
		
		
		XSLFAutoShape shp = (XSLFAutoShape) slide1.getPlaceholder(1);
		Rectangle2D anchor = shp.getAnchor();
		shp.clearText();
		XSLFPictureShape picture = slide1.createPicture(idx);
        slide1.removeShape(shp);

        picture.setAnchor(anchor);  
		//shp.drawContent((Graphics2D) myPicture.getGraphics());
		
		XSLFTextShape title2 =  slide1.getPlaceholder(0);
		title2.setText("Model 2");
		*/
		
		//-----slide 2-----
		//**XSLFGroupShape group1 = slide2.createGroup(); 
		
		//**picture = group1.createPicture(idx); 
				
		///////////////
		FileOutputStream out = new FileOutputStream(file);
		ppt.write(out);
		System.out.println("PowerPoint with image created!");

		out.close();

		showAvailableSlideLayouts();
	}

	private static void createSlideTitle(XSLFSlide slide, String string) {
		//ADD TEXT FOR TITLE
		XSLFTextBox shape = slide.createTextBox();
		shape.setAnchor(new Rectangle(250, 5, 150, 25));
		XSLFTextParagraph p = shape.addNewTextParagraph();
		p.setTextAlign(TextAlign.CENTER);

		XSLFTextRun r1 = p.addNewTextRun();
		r1.setText("Model 1");
		r1.setFontColor(Color.BLACK);
		r1.setFontSize(24.);
		
	}

	//CHART
	private static XYZDataset createDataset() {
		DefaultXYZDataset dataset = new DefaultXYZDataset();

		dataset.addSeries("GI",new double[][] { { 3,5 }, // X-Value
						{ 50,60 }, // Y-Value
						{ 20,30 } // Z-Value
				});
		dataset.addSeries("GT", new double[][] { { 1 }, { 75 }, { 20 } });
		
		return dataset;
	}
	
	private static JFreeChart createBubbleChart() {
		// Create dataset
		XYZDataset dataset = createDataset();

		// Create chart
		JFreeChart chart = ChartFactory.createBubbleChart("Bubble demo", "X-Values", "Y-Values", dataset);

		// Set range for X-Axis
		XYPlot plot = chart.getXYPlot();
		NumberAxis domain = (NumberAxis) plot.getDomainAxis();
		domain.setRange(0, 5);

		// Set range for Y-Axis
		NumberAxis range = (NumberAxis) plot.getRangeAxis();
		range.setRange(0, 100);
		

		return chart;
	}
	
	//PIE
	private static PieDataset createDatasetPie( Double target, Double currentVal) {
	      DefaultPieDataset dataset = new DefaultPieDataset( );
	      dataset.setValue( "Robustness", currentVal);  
	      dataset.setValue( "Target", Math.abs(target-currentVal));   
	      return dataset;         
	}
	
	private static JFreeChart createPieChart(PieDataset dataset, String chartName ) {
	      JFreeChart chart = ChartFactory.createPieChart(      
	         chartName,   // chart title 
	         dataset,          // data    
	         true,             // include legend   
	         true, 
	         false);

	      chart.getPlot().setBackgroundPaint(Color.WHITE);
	      return chart;
	  }

	private static JFreeChart createRingChart(PieDataset pieDataset, String chartName ) {
	      JFreeChart chart = ChartFactory.createRingChart(     
	         chartName,   // chart title 
	         pieDataset,          // data    
	         true,             // include legend   
	         true,      //tooltip
	         false);

	      chart.getPlot().setBackgroundPaint(Color.WHITE);
	      return chart;
	  }
	
	private static CategoryDataset createProgressBarDataset() {
		
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();		 
		dataset.addValue(4.0, "Target", "");
		dataset.addValue(3.6, "Index", "");
		return dataset;		
	}
	
	private static JFreeChart createBarChart(CategoryDataset dataset, String chartName ) {
		 JFreeChart chart=ChartFactory.createBarChart(
			        "Robustness", //Chart Title
			        "Robustness vs Target", // Category axis
			        "Target", // Value axis 
			        dataset,
			        PlotOrientation.VERTICAL,
			        true,true,false
			       );


         // NOW DO SOME OPTIONAL CUSTOMISATION OF THE CHART...

         // set the background color for the chart...
         chart.setBackgroundPaint(Color.WHITE);

         // get a reference to the plot for further customisation...
         final CategoryPlot plot = chart.getCategoryPlot();
            plot.setRangeAxisLocation(AxisLocation.BOTTOM_OR_LEFT);
            
         // change the auto tick unit selection to integer units only...
         final NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
         
         rangeAxis.setRange(0.0, 4.8);
         rangeAxis.setStandardTickUnits(NumberAxis.createStandardTickUnits());
            
         // OPTIONAL CUSTOMISATION COMPLETED
            
         return chart;
	}
	
	private static JFreeChart createStackedBarChart(CategoryDataset dataset, String chartName ) {
		final JFreeChart chart = ChartFactory.createStackedBarChart(
				  chartName, 
				  "", 
				  chartName + " vs Target",
				  dataset, 
				  PlotOrientation.HORIZONTAL, 
				  true, 
				  true, 
				  false);

				  chart.setBackgroundPaint(new Color(255, 255, 255));

				  CategoryPlot plot = chart.getCategoryPlot();
				  plot.getRenderer().setSeriesPaint(0, new Color(248, 82, 47));
				  plot.getRenderer().setSeriesPaint(1, new Color(240, 230, 19));
				  plot.getRenderer().setSeriesPaint(2, new Color(19, 201, 7));

				  return chart;
	}
	
	private static CategoryDataset createStackedBarDataset(double index, double target, double threshold) {
		/*  double[][] data = new double[][]{
		  {210},
		  {200},
		  };
		  CategoryDataset ds =  DatasetUtilities.createCategoryDataset(
		  "Team ", "Match", data);*/
		
		/*double index = 3.6;
		double target = 4.8;
		double threshold = 4.0;*/
		
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();	
		
		dataset.addValue(index, "Index", "");
		dataset.addValue(Math.abs(threshold-index), "Treshold", "");
		dataset.addValue(target - Math.max(index,threshold), "Target", "");
		
		return dataset;	
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

	private static void createTableOnSlide(File file, XMLSlideShow ppt, XSLFSlide slide) throws IOException {

		// Add table to ppt on the same page
		
		int numColumns = 3;
		int numRows = demoData().size();
				
		XSLFTable tbl = slide.createTable();
		
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
		
			DemoDataObj dataRow = demoData().get(rownum);
			
			XSLFTableCell cell = tr.addCell();
			XSLFTextParagraph p = cell.addNewTextParagraph();
			XSLFTextRun r = p.addNewTextRun();
			r.setText(dataRow.getCompany());
			r.setFontSize(10);
			p.addLineBreak();
			
			r = p.addNewTextRun();			
			r.setText("Robustness");
			r.setFontSize(6);
						
			
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

	private static void createPresentationWithImage(File image, File file, XMLSlideShow ppt, XSLFSlide slide, Rectangle rectangle)
			throws IOException {
		// XMLSlideShow ppt = new XMLSlideShow();
		// XSLFSlide slide = ppt.createSlide();
		
		byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
		int idx = ppt.addPicture(picture, XSLFPictureData.PICTURE_TYPE_PNG);
		XSLFPictureShape pic = slide.createPicture(idx);
		pic.setAnchor(rectangle);
		

		// File file = new File("PowerPointBubble.pptx");
		/*
		 * FileOutputStream out = new FileOutputStream(file); ppt.write(out);
		 * System.out.println("PowerPoint with image created!");
		 * 
		 * out.close();
		 */
	}

	public static void showAvailableSlideLayouts() {
		// create an empty presentation
		XMLSlideShow ppt = new XMLSlideShow();
		System.out.println("Available slide layouts:");

		// getting the list of all slide masters
		for (XSLFSlideMaster master : ppt.getSlideMasters()) {

			// getting the list of the layouts in each slide master
			for (XSLFSlideLayout layout : master.getSlideLayouts()) {

				// getting the list of available slides
				System.out.println(layout.getType());
			}
		}
	}
	
}
