package sugaryo.poi_feasibility.core;

import sugaryo.poi_feasibility.utility.ExcelWrapper;

public class Report {
	
	public void output() {
		
		// 取り敢えず新規にブックを作って保存する。
		try ( var excel = new ExcelWrapper() ) { 
			excel.cell( 0, 0 ).value( "test" );
			excel.save( "C:/test/poi/test.xlsx" );
		}
		// 検査例外は RuntimeException でラップしてスローする。
		catch ( Exception ex ) {
			throw new RuntimeException( ex );
		}
	}
}
