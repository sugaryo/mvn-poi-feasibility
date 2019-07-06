package sugaryo.poi_feasibility.utility;

import static sugaryo.poi_feasibility.utility.PoiUtil.serialize;

import java.util.ArrayList;
import java.util.List;

import static sugaryo.poi_feasibility.utility.PoiUtil.output;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelWrapper implements AutoCloseable {
	
	
	private final XSSFWorkbook book;
	
	private XSSFSheet current;
	
	public class CellContext {
		
		private final XSSFCell xcell;
		
		private CellContext(XSSFCell xcell) {
			this.xcell = xcell;
		}
		
		
		public CellContext value(String str) {
			
			this.xcell.setCellValue( str );
			return this;
		}
		
		// TODO：他にもユーティリティを追加。
	}
	
	public class RangeContext {
		
		private final XSSFCell xcell1;
		private final XSSFCell xcell2;
		
		public RangeContext(XSSFCell xcell1, XSSFCell xcell2) {
			this.xcell1 = xcell1;
			this.xcell2 = xcell2;
		}

		
		public RangeContext clearRows() {
			
			final XSSFSheet sheet = this.xcell1.getSheet();
			final int row1 = this.xcell1.getRowIndex();
			final int row2 = this.xcell2.getRowIndex();
			final int n = row2 - row1 + 1; // 植木算
			
			
			// ■先に範囲内に含まれる結合セルを解除。
			List<CellRangeAddress> regions = sheet.getMergedRegions();
			List<Integer> unmerges = new ArrayList<>();
			for ( int i = 0; i < regions.size(); i++ ) {
				var region = regions.get(i);
				
				// 行範囲内に内包している結合セルを解除。
				// （引っ掛かってるやつも対象にして良い気がするけど、そうなるケースってそもそも名前定義の範囲が良くない）
				if ( row1 <= region.getFirstRow() && region.getLastRow() <= row2 ) {
					unmerges.add(i);
				}
			}
			sheet.removeMergedRegions( unmerges );
			
			
			// ■範囲内の行データを削除。（Rowが消えるのでRowオブジェクトは再度createする）
			for ( int i = 0; i < n; i++ ) {
				int row = row1 + i;
				var xrow = sheet.getRow( row );
				sheet.removeRow( xrow );
				sheet.createRow( row );
			}
			
			return this;
		}
		public RangeContext hideRows() {

			final XSSFSheet sheet = this.xcell1.getSheet();
			final int row1 = this.xcell1.getRowIndex();
			final int row2 = this.xcell2.getRowIndex();
			final int n = row2 - row1 + 1; // 植木算
			
			for ( int i = 0; i < n; i++ ) {
				int row = row1 + i;
				var xrow = sheet.getRow( row );
				xrow.setZeroHeight( true );
			}
			
			return this;
		}
		
		public RangeContext insertRows() {
			return this.insertRows( false );
		}
		public RangeContext insertRows( boolean withCopy ) {
			throw new RuntimeException( "まだつくってないよ" ); // TODO：実装
		}
		
		public RangeContext deleteRows() {

			final XSSFSheet sheet = this.xcell1.getSheet();
			final int row1 = this.xcell1.getRowIndex();
			final int row2 = this.xcell2.getRowIndex();
			final int n = row2 - row1 + 1; // 植木算
			
			
			// ■先に範囲内に含まれる結合セルを解除。
			var regions = sheet.getMergedRegions();
			List<Integer> unmerges = new ArrayList<>();
			for ( int i = 0; i < regions.size(); i++ ) {
				var region = regions.get(i);
				
				// 行範囲内に内包している結合セルを解除。
				// （引っ掛かってるやつも対象にして良い気がするけど、そうなるケースってそもそも名前定義の範囲が良くない）
				if ( row1 <= region.getFirstRow() && region.getLastRow() <= row2 ) {
					unmerges.add(i);
				}
			}
			sheet.removeMergedRegions( unmerges );
			
			
			// ■範囲内の行データを削除。
			for ( int i = 0; i < n; i++ ) {
				int row = row1 + i;
				var xrow = sheet.getRow( row );
				sheet.removeRow( xrow );
			}
			
			// ■消して空いた領域に行シフト。
			sheet.shiftRows( row2 + 1, sheet.getLastRowNum(), -n );

			return this;
		}
	}
	
	
	
	public ExcelWrapper(XSSFWorkbook book) {
		this.book = book;
		this.current = this.book.getSheetAt( 0 ); // 取り敢えずデフォルトでは先頭シート選択。
	}
	
	public ExcelWrapper() {
		this.book = new XSSFWorkbook();
		this.current = this.book.createSheet();
	}
	
	
	public ExcelWrapper sheet(int index) {
		this.current = this.book.getSheetAt( index );
		return this;
	}
	
	public ExcelWrapper sheet(String name) {
		this.current = this.book.getSheet( name );
		return this;
	}
	
	public CellContext cell(final int row, final int col) {
		
		var xrow = this.current.getRow( row );
		if ( null == xrow ) xrow = this.current.createRow( row );
		
		var xcell = xrow.getCell( col );
		if ( null == xcell ) xcell = xrow.createCell( col );
		
		return new CellContext( xcell );
	}
	
	
	public CellContext cell(String name) {
		return this.cell( name, true );
	}
	
	public CellContext cell(String name, boolean withSheetSelect) {
		
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
		// ■シート選択しない場合：
		else {
			var sheet = this.book.getSheet( sheetname );
			return new CellContext( sheet.getRow( row ).getCell( col ) );
		}
	}
	
	
	public RangeContext range(String name) {
		return this.range( name, true );
	}
	
	public RangeContext range(String name, boolean withSheetSelect) {
		
		// 名前定義から [SheetName, Row, Column] を取得。
		var xname = this.book.getName( name );
		var ref = new AreaReference( xname.getRefersToFormula(), SpreadsheetVersion.EXCEL2007 );
		
		
		final String sheetname = xname.getSheetName();
		
		// ■シート選択する場合：
		if ( withSheetSelect ) {
			
			this.sheet( sheetname );
		}
		
		return this.range( this.book, ref );
	}

	private RangeContext range(XSSFWorkbook book, AreaReference ref) {
		var xcell1 = xcell( book, ref.getFirstCell() );
		var xcell2 = xcell( book, ref.getLastCell() );
		return new RangeContext( xcell1, xcell2 );
	}
	
	private static XSSFCell xcell(XSSFWorkbook book, CellReference ref) {
		var sheet = book.getSheet( ref.getSheetName() );
		
		
		// TODO：厳密指定するオプションを追加したほうが良いか？
		var xrow = sheet.getRow( ref.isRowAbsolute() ? ref.getRow() : 0 );
		var xcell = xrow.getCell( ref.isColAbsolute() ? ref.getCol() : 0 );
		
		return xcell;
	}
	
	
	public byte[] binary() {
		return serialize( this.book );
	}
	
	public void save(String path) {
		output( this.book, path );
	}
	
	/** @inherit */
	@Override
	public void close() throws Exception {
		
		this.book.close();
	}
}
