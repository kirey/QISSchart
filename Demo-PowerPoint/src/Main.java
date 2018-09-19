import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import javax.imageio.ImageIO;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;

public class Main {


	public static void main(String[] args) throws IOException {
		export();
		System.out.println("GOTOVO");
	}

	public static void export() throws IOException {
		BufferedImage objBufferedImage = createChart(createDataset()).createBufferedImage(600, 800);
		ByteArrayOutputStream bas = new ByteArrayOutputStream();
        try {
            ImageIO.write(objBufferedImage, "png", bas);
        } catch (IOException e) {
            e.printStackTrace();
        }

        byte[] byteArray=bas.toByteArray();
        InputStream in = new ByteArrayInputStream(byteArray);
        BufferedImage image;
		try {
			image = ImageIO.read(in);
			File outputfile = new File("image.png");
	        ImageIO.write(image, "png", outputfile);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
	}

	private static JFreeChart createChart(PieDataset dataset) {
		JFreeChart chart = ChartFactory.createPieChart("Mobile Sales", // chart title
				dataset, // data
				true, // include legend
				true, false);

		return chart;
	}

	private static PieDataset createDataset() {
		DefaultPieDataset dataset = new DefaultPieDataset();
		dataset.setValue("IPhone 5s", new Double(20));
		dataset.setValue("SamSung Grand", new Double(20));
		dataset.setValue("MotoG", new Double(40));
		dataset.setValue("Nokia Lumia", new Double(10));
		return dataset;
	}

}
