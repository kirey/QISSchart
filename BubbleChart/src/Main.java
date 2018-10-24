import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
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
import org.jfree.chart.plot.CenterTextMode;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.PiePlot3D;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.RingPlot;
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
		
		File pptFileModel1 = new File("Model1.pptx");
		Model1.generateModel1(pptFileModel1);
		
		File pptFileModel2 = new File("Model2.pptx");
		Model2.generateModel(pptFileModel2);
		
		File pptFileModel3 = new File("Model3.pptx");
		Model3.generateModel(pptFileModel3);
		
		File pptFileModel4 = new File("Model4.pptx");
		//Model4.generateModel(pptFileModel4);
		
	/*
		//////TEST///// inserting image to layout content
		
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
		title2.setText("Model 2");		 */
	
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
