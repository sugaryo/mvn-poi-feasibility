package sugaryo.poi_feasibility;

import sugaryo.poi_feasibility.core.Report;

public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        
        try {
        	// 実際にはWebシステムか何かで
        	// ReportFactory的な物を通して使い、
        	// バイナリをoctet-streamで返すのを想定。
			new Report().output();
			
			System.out.println( "ok" );
		} catch ( Exception ex ) {
			System.err.println( ex );
		}
    }
}
