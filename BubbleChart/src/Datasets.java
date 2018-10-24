import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;
import org.jfree.data.xy.DefaultXYZDataset;
import org.jfree.data.xy.XYZDataset;

public class Datasets {

	//CHART
	public static XYZDataset createDataset() {
		DefaultXYZDataset dataset = new DefaultXYZDataset();

		dataset.addSeries("GI",new double[][] { { 3 }, // X-Value
							{ 50 }, // Y-Value
							{ 20 } // Z-Value
					});
		dataset.addSeries("GT", new double[][] { { 1 }, { 75 }, { 20 } });
			
		return dataset;
	}
		
		
		
	//PIE
	public static PieDataset createDatasetPie( Double target, Double currentVal) {
	      DefaultPieDataset dataset = new DefaultPieDataset( );
	      dataset.setValue( "Robustness", currentVal);  
	      dataset.setValue( "Target", Math.abs(target-currentVal));   
	      return dataset;         
	}
			
	public static CategoryDataset createProgressBarDataset() {
			
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();		 
		dataset.addValue(4.0, "Target", "");
		dataset.addValue(3.6, "Index", "");
		return dataset;		
	}
	
	public static CategoryDataset createStackedBarDataset(double index, double target, double threshold) {
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
	
	public static CategoryDataset createStackedBarDataset(double automatic, double manual) {
		
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();	
		
		dataset.addValue(automatic, "Automatic", "");
		dataset.addValue(manual, "Manual", "");
				
		return dataset;	
	}	
	
}
