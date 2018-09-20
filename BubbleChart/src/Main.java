import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.plot.XYPlot;
import org.jfree.data.xy.DefaultXYZDataset;
import org.jfree.data.xy.XYZDataset;

public class Main {

	public static void main(String[] args) throws IOException {

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

		File bubbleChart = new File("BubbleChart.png");
		int width = 560; /* Width of the image */
		int height = 370; /* Height of the image */
		ChartUtilities.saveChartAsPNG(bubbleChart, chart, width, height);
		System.out.println("Image rendered!");
		createPresentationWithImage(bubbleChart);
	}

	private static XYZDataset createDataset() {
		DefaultXYZDataset dataset = new DefaultXYZDataset();

		dataset.addSeries("GI", new double[][] { { 3 }, // X-Value
				{ 100 }, // Y-Value
				{ 20 } // Z-Value
		});
		dataset.addSeries("GT", new double[][] { { 1 }, { 100 }, { 20 } });

		return dataset;
	}

	private static void createPresentationWithImage(File image) throws IOException {
		XMLSlideShow ppt = new XMLSlideShow();
		XSLFSlide slide = ppt.createSlide();

		byte[] picture = IOUtils.toByteArray(new FileInputStream(image));
		int idx = ppt.addPicture(picture, XSLFPictureData.PICTURE_TYPE_PNG);
	    XSLFPictureShape pic = slide.createPicture(idx);
		File file = new File("PowerPointBubble.pptx");
		FileOutputStream out = new FileOutputStream(file);
		ppt.write(out);
		System.out.println("PowerPoint with image created!");
		out.close();
	}

}
