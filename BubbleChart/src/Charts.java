import java.awt.Color;
import java.awt.Font;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.AxisLocation;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.CenterTextMode;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.PiePlot3D;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.RingPlot;
import org.jfree.chart.plot.XYPlot;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.general.PieDataset;
import org.jfree.data.xy.XYZDataset;

public class Charts {
	
	public static JFreeChart createBubbleChart(XYZDataset dataset, String chartName, String xAxisName, String yAxisName) {
		// Create dataset
		//XYZDataset dataset = createDataset();

		// Create chart
		JFreeChart chart = ChartFactory.createBubbleChart(chartName, xAxisName, yAxisName, dataset);

		// Set range for X-Axis
		XYPlot plot = chart.getXYPlot();
		NumberAxis domain = (NumberAxis) plot.getDomainAxis();
		domain.setRange(0, 5);

		// Set range for Y-Axis
		NumberAxis range = (NumberAxis) plot.getRangeAxis();
		range.setRange(0, 100);
		
		chart.getPlot().setBackgroundPaint(Color.WHITE);
		plot.getRenderer().setSeriesPaint(0, new Color(180, 180, 200));
		plot.getRenderer().setSeriesPaint(1, new Color(253, 192, 179));
		plot.getRenderer().setSeriesPaint(2, new Color(174, 231, 231));

		return chart;
	}

	public static JFreeChart createPieChart(PieDataset dataset, String chartName ) {
	      JFreeChart chart = ChartFactory.createPieChart(      
	         chartName,   // chart title 
	         dataset,          // data    
	         true,             // include legend   
	         true, 
	         false);

	      chart.getPlot().setBackgroundPaint(Color.WHITE);
	      chart.getPlot().setOutlineVisible(false);
	      chart.getPlot().setNoDataMessage("No data to display");
	      ((PiePlot) chart.getPlot()).setLabelGenerator(null);     // not showing labels on pie
	      ((PiePlot) chart.getPlot()).setLabelFont(new Font("Arial", Font.PLAIN, 10));
	      return chart;
	  }
	
	public static JFreeChart createPieChart3D(PieDataset dataset, String chartName ) {
	      JFreeChart chart = ChartFactory.createPieChart3D(      
	         chartName,   // chart title 
	         dataset,          // data    
	         true,             // include legend   
	         true, 
	         false);

	      final PiePlot3D plot = ( PiePlot3D ) chart.getPlot( );             
	      plot.setStartAngle( 270 );             
	      plot.setForegroundAlpha( 0.65f );             
	      plot.setInteriorGap( 0.02 );   
	      plot.setLabelFont(new Font("Arial", Font.PLAIN, 10));
	      
	      plot.setBackgroundPaint(Color.WHITE);
	      plot.setOutlineVisible(false);
	      plot.setNoDataMessage("No data to display");
	      plot.setLabelGenerator(null);
	      return chart;
	  }

	public static JFreeChart createRingChart(PieDataset pieDataset, String chartName, String type ) {
	      JFreeChart chart = ChartFactory.createRingChart(     
	         chartName,   // chart title 
	         pieDataset,          // data    
	         false,             // include legend or not  
	         true,      //tooltip
	         false);

	      RingPlot plot = (RingPlot) chart.getPlot();
	      plot.setBackgroundPaint(Color.WHITE);
	      plot.setOutlineVisible(false);  //removes border around chart
	      
	      //COLOR OF SECTIONS
	      if(type.equals("TYPE1")){ 
	    	  plot.setSectionPaint(0, new Color(153,153,153));
		      plot.setSectionPaint(1, new Color(248,248,248));
			  
		  }else if(type.equals("TYPE2")){
			  plot.setSectionPaint(2, new Color(248, 82, 47));
		      plot.setSectionPaint(1, new Color(240, 230, 19));
		      plot.setSectionPaint(0, new Color(19, 201, 7));	
		      
		      plot.setSectionDepth(0.5);
		  }
	            
	      //SHADOWS
	      plot.setShadowXOffset(0);   //removes shadows
	      plot.setShadowYOffset(0);  //removes shadows
	      
	      //TEXT IN CENTER
	      if(type.equals("TYPE1")){
		      plot.setCenterText(pieDataset.getValue(0).toString()); 
		      plot.setCenterTextFont(new Font("Arial",Font.BOLD, 25));
		      plot.setCenterTextColor(Color.BLACK);
		      plot.setCenterTextMode(CenterTextMode.FIXED);
	      }
	      plot.setLabelGenerator(null);
	      
	      return chart;
	  }
	
	public static JFreeChart createBarChart(CategoryDataset dataset, String chartName ) {
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
	
	public static JFreeChart createStackedBarChart(CategoryDataset dataset, String chartName ) {
		final JFreeChart chart = ChartFactory.createStackedBarChart(
				  "", 
				  "", 
				  "",
				  dataset, 
				  PlotOrientation.HORIZONTAL, 
				  true, 
				  true, 
				  false);

				  chart.setBackgroundPaint(new Color(255, 255, 255));
				  chart.removeLegend();
				  
				
				  CategoryPlot plot = chart.getCategoryPlot();
				  plot.getRenderer().setSeriesPaint(0, new Color(248, 82, 47));
				  plot.getRenderer().setSeriesPaint(1, new Color(240, 230, 19));
				  plot.getRenderer().setSeriesPaint(2, new Color(19, 201, 7));
				
				  return chart;
	}
	
	public static JFreeChart stackedBarChartNoAxis(CategoryDataset dataset, String chartName, String type) {
		final JFreeChart chart = ChartFactory.createStackedBarChart(
				  "", 
				  "", 
				  "",
				  dataset, 
				  PlotOrientation.HORIZONTAL, 
				  true, 
				  true, 
				  false);

				  chart.setBackgroundPaint(new Color(255, 255, 255));
				  chart.removeLegend();
				  
				  CategoryPlot plot = chart.getCategoryPlot();
				  
				  if(type.equals("TYPE1")){ 
					  plot.getRenderer().setSeriesPaint(0, new Color(43, 120, 228));
					  plot.getRenderer().setSeriesPaint(1, new Color(153, 0, 255));	
					  
				  }else if(type.equals("TYPE2")){
					  plot.getRenderer().setSeriesPaint(0, new Color(248, 82, 47));
					  plot.getRenderer().setSeriesPaint(1, new Color(240, 230, 19));
					  plot.getRenderer().setSeriesPaint(2, new Color(19, 201, 7));				  
				  }
				  
				  //removes range axis
				  ValueAxis range = plot.getRangeAxis();
				  range.setVisible(false);
				
				  return chart;
	}
	
}
