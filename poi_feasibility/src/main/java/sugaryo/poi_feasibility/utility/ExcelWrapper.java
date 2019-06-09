package sugaryo.poi_feasibility.utility;

import static sugaryo.poi_feasibility.utility.PoiUtil.serialize;
import static sugaryo.poi_feasibility.utility.PoiUtil.output;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelWrapper implements AutoCloseable {
	
	
	private final XSSFWorkbook book;
	
	private XSSFSheet current;
	
	public class CellContext {
		
		private final XSSFCell xcell;
		
		private CellContext( XSSFCell xcell ) {
			this.xcell = xcell;
		}
		
		
		public CellContext value( String str ) {
			this.xcell.setCellValue( str );
			return this;
		}
		
		// TODO：他にもユーティリティを追加。
	}
	
	
	
	public ExcelWrapper( XSSFWorkbook book ) {
		this.book = book;
		this.current = this.book.getSheetAt( 0 ); // 取り敢えずデフォルトでは先頭シート選択。
	}
	
	public ExcelWrapper() {
		this.book = new XSSFWorkbook();
		this.current = this.book.createSheet();
	}
	
	
	public ExcelWrapper sheet( int index ) {
		this.current = this.book.getSheetAt( index );
		return this;
	}
	
	public ExcelWrapper sheet( String name ) {
		this.current = this.book.getSheet( name );
		return this;
	}
	
	public CellContext cell( final int row, final int col ) {
		
		var xrow = this.current.getRow( row );
		if ( null == xrow ) xrow = this.current.createRow( row );
		
		var xcell = xrow.getCell( col );
		if ( null == xcell ) xcell = xrow.createCell( col );
		
		return new CellContext( xcell );
	}
	
	public CellContext cell( String name ) {
		return this.cell( name, true );
	}
	
	public CellContext cell( String name, boolean withSheetSelect ) {
		
		// 名前定義から [SheetName, Row, Column] を取得。
		var xname = this.book.getName( name );
		var ref = new CellReference( xname.getRefersToFormula() );
		
		final String sheetname = ref.getSheetName();
		final int row = ref.getRow();
		final int col = ref.getCol();
		
		
		// ■シート選択する場合：
		if ( withSheetSelect ) {
			
			return this.sheet( sheetname ).cell( row, col );
			
		}
		// ■シート選択肢ない場合：
		else {
			var sheet = this.book.getSheet( sheetname );
			return new CellContext( sheet.getRow( row ).getCell( col ) );
		}
	}
	
	
	public byte[] binary() {
		return serialize( this.book );
	}
	
	public void save( String path ) {
		output( this.book, path );
	}
	
	/** @inherit */
	@Override
	public void close() throws Exception {
		//close は v3.10.1 時点では存在しない。
		//this.book.close();
	}
}
