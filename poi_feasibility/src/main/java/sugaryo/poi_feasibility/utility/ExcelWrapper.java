package sugaryo.poi_feasibility.utility;

import static sugaryo.poi_feasibility.utility.PoiUtil.serialize;
import static sugaryo.poi_feasibility.utility.PoiUtil.output;
import static sugaryo.poi_feasibility.utility.PoiUtil.poiShiftRows;
import static sugaryo.poi_feasibility.utility.PoiUtil.poiHideRows;

import java.util.ArrayList;
import java.util.List;


import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelWrapper implements AutoCloseable {

	private static final CellCopyPolicy DEFAULT_COPY_POLICY = new CellCopyPolicy.Builder()
			.rowHeight( true )
			.cellStyle( true )
			.cellValue( true )
			.mergedRegions( true )
			.build(); 
	
	private final XSSFWorkbook book;
	
	private XSSFSheet current;
	
	public static class CellContext {

		private final XSSFSheet sheet;
		private final XSSFCell xcell;
		
		private CellContext(XSSFCell xcell) {
			this.xcell = xcell;
			this.sheet = xcell.getSheet();
		}
		
		public int row() {
			return this.xcell.getRowIndex();
		}
		public int col() {
			return this.xcell.getColumnIndex();
		}
		
		
		public CellContext value(String str) {
			this.xcell.setCellValue( str );
			return this;
		}
		
		
		public CellContext cellBreak() {
			return this.rowBreak().colBreak();
		}
		
		public CellContext rowBreak() {
			this.sheet.setRowBreak( this.row() - 1 ); // Excelの操作感に合わせるために１ずらす。
			return this;
		}
		
		public CellContext colBreak() {
			this.sheet.setColumnBreak( this.col() - 1 ); // Excelの操作感に合わせるために１ずらす。
			return this;
		}
	}
	
	public static class RangeContext {

		private final XSSFSheet sheet;
		private final CellReference ref1; // 基準側：rectangleの左上相当セル
		private final CellReference ref2; // 終点側：rectangleの右下相当セル
		
		// ctor
		private RangeContext(XSSFSheet sheet, CellReference ref1, CellReference ref2) {
			this.sheet = sheet;
			this.ref1 = ref1;
			this.ref2 = ref2;
		}
		
		
		// readonly-prop;
		
		public int rows() {
			final int row1 = this.ref1.getRow();
			final int row2 = this.ref2.getRow();
			return row2 - row1 + 1; // 植木算
		}
		public int cols() {
			final int col1 = this.ref1.getCol();
			final int col2 = this.ref2.getCol();
			return col2 - col1 + 1; // 植木算
		}
		
		public int top() {
			return this.ref1.getRow();
		}
		public int bottom() {
			return this.ref2.getRow();
		}
		public int left() {
			return this.ref1.getCol();
		}
		public int right() {
			return this.ref2.getCol();
		}
		
		public boolean isSingleRow() {
			return this.top() == this.bottom();
		}
		public boolean isMultipleRow() {
			return this.top() != this.bottom();
		}
		
		
		// page-break-utility;

		public RangeContext topBreak() {
			this.sheet.setRowBreak( this.top() - 1 ); // top位置でPageBreak;
			return this;
		}
		
		public RangeContext bottomBreak() {
			this.sheet.setRowBreak( this.bottom() ); // bottom直下でPageBreak;
			return this;
		}
		
		public RangeContext leftBreak() {
			this.sheet.setColumnBreak( this.ref1.getCol() - 1 ); // left位置でPageBreak;
			return this;
		}
		
		public RangeContext rightBreak() {
			this.sheet.setColumnBreak( this.ref2.getCol() ); // rightの右側でPageBreak;
			return this;
		}
		
		
		// row-level-utility
		
		public RangeContext clearRows() {

			final int row1 = this.ref1.getRow();
			final int row2 = this.ref2.getRow();
			final int rows = row2 - row1 + 1; // 植木算
			
			
			// ■先に範囲内に含まれる結合セルを解除。
			List<CellRangeAddress> regions = this.sheet.getMergedRegions();
			List<Integer> unmerges = new ArrayList<>();
			for ( int i = 0; i < regions.size(); i++ ) {
				CellRangeAddress region = regions.get( i );
				
				// 行範囲内に内包している結合セルを解除。
				// （引っ掛かってるやつも対象にして良い気がするけど、そうなるケースってそもそも名前定義の範囲が良くない）
				if ( row1 <= region.getFirstRow() && region.getLastRow() <= row2 ) {
					unmerges.add( i );
				}
			}
			this.sheet.removeMergedRegions( unmerges );
			
			
			// ■範囲内の行データを削除。（Rowが消えるのでRowオブジェクトは再度createする）
			for ( int i = 0; i < rows; i++ ) {
				int row = row1 + i;
				var xrow = this.sheet.getRow( row );
				// 行オブジェクトがあれば remove する。
				// ※ テンプレート上でデータがない場合は行オブジェクト自体いない事もある
				if ( null != xrow ) {
					this.sheet.removeRow( xrow );
				}
				this.sheet.createRow( row );
			}
			
			return this;
		}
		public RangeContext hideRows() {

			final int row1 = this.ref1.getRow();
			final int row2 = this.ref2.getRow();
			final int rows = row2 - row1 + 1; // 植木算
			
			poiHideRows( this.sheet, row1, rows, false );
			
			return this;
		}
		
		public RangeContext insertRows( final int count ) {

			final int row1 = this.ref1.getRow();
			final int row2 = this.ref2.getRow();
			final int rows = row2 - row1 + 1; // 植木算
			
			// この領域の上に、指定個数ぶんの空行を作る（シフトして場所を開けるだけ）
			poiShiftRows( this.sheet, row1, rows * count );
			
			return this;
		}
		
		
		public RangeContext copyRows( final int count ) {
			return this.copyRows( count, DEFAULT_COPY_POLICY );
		}
		public RangeContext copyRows( final int count, final CellCopyPolicy policy ) {

			final int row1 = this.ref1.getRow();
			final int row2 = this.ref2.getRow();
			final int rows = row2 - row1 + 1; // 植木算
			
			
			// ■シフト処理：
			poiShiftRows( this.sheet, row2 + 1, rows * count );
			
			// ■コピー処理：
			for ( int n = 0; n < count; n++ ) {
				
				final int shift = n * rows;
				final int dest = row2 + 1 + shift;
				
				this.sheet.copyRows( row1, row2, dest, policy );
			}
			
			return this;
		}
		
		public RangeContext deleteRows() {

			final int row1 = this.ref1.getRow();
			final int row2 = this.ref2.getRow();
			final int rows = row2 - row1 + 1; // 植木算
			
			
			// ■先に範囲内に含まれる結合セルを解除。
			List<CellRangeAddress> regions = this.sheet.getMergedRegions();
			List<Integer> unmerges = new ArrayList<>();
			for ( int i = 0; i < regions.size(); i++ ) {
				CellRangeAddress region = regions.get( i );
				
				// 行範囲内に内包している結合セルを解除。
				// （引っ掛かってるやつも対象にして良い気がするけど、そうなるケースってそもそも名前定義の範囲が良くない）
				if ( row1 <= region.getFirstRow() && region.getLastRow() <= row2 ) {
					unmerges.add( i );
				}
			}
			this.sheet.removeMergedRegions( unmerges );
			
			
			// ■範囲内の行データを削除。
			for ( int i = 0; i < rows; i++ ) {
				int row = row1 + i;
				var xrow = this.sheet.getRow( row );
				
				// 行オブジェクトがあれば remove する。
				// ※ テンプレート上でデータがない場合は行オブジェクト自体いない事もある
				if ( null != xrow ) {
					this.sheet.removeRow( xrow );
				}
			}
			
			// ■消して空いた領域に行シフト。
			poiShiftRows( this.sheet, row2 + 1, -rows );

			return this;
		}
		
		
		// immutable-class-utility;
		
		/**
		 * 領域の平行移動.
		 * 
		 * @param dr 行変位量
		 * @param dc 列変位量
		 * @return 変位量を適合した新しい {@link RangeContext} オブジェクト
		 */
		public RangeContext move(int dr, int dc) {
			if ( 0 == dr && 0 == dc ) return this;
			
			// 基点・終点ともに同僚移動させる（単純な平行移動）
			var next1 = new CellReference( this.ref1.getRow() + dr, this.ref1.getCol() + dc );
			var next2 = new CellReference( this.ref2.getRow() + dr, this.ref2.getCol() + dc );
			return new RangeContext( this.sheet, next1, next2 );
		}
		
		/**
		 * 領域の拡張.
		 * 
		 * @param dr 行変位量
		 * @param dc 列変位量
		 * @return 変位量を適合した新しい {@link RangeContext} オブジェクト
		 */
		public RangeContext fat(int dr, int dc) {
			if ( 0 == dr && 0 == dc ) return this;
			
			// 基準側は移動させず、終点側を移動させる。
			var next1 = new CellReference( this.ref1.getRow() + 0, this.ref1.getCol() + 0 );
			var next2 = new CellReference( this.ref2.getRow() + dr, this.ref2.getCol() + dc );
			return new RangeContext( this.sheet, next1, next2 );
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
	
	
	public ExcelWrapper header(String left, String center, String right) {
		var header = this.current.getHeader();
		if ( null != left ) header.setLeft( left );
		if ( null != center ) header.setCenter( center );
		if ( null != right ) header.setRight( right );
		return this;
	}
	public ExcelWrapper footer(String left, String center, String right) {
		var footer = this.current.getFooter();
		if ( null != left ) footer.setLeft( left );
		if ( null != center ) footer.setCenter( center );
		if ( null != right ) footer.setRight( right );
		return this;
	}
	
	
	public boolean exists(String name) {
		return null != this.book.getName( name );
	}
	
	
	public CellContext cell(final int row, final int col) {
		
		var xcell = this.xssfCell( row, col );
		return new CellContext( xcell );
	}
	
	
	public CellContext cell(String name) {
		
		// 名前定義から [SheetName, Row, Column] を取得。
		var xname = this.book.getName( name );
		
		// TODO：指定した名前定義が無かった場合の動作オプション（例外 or NOP代替）を考える。
		
		var ref = new CellReference( xname.getRefersToFormula() );
		
		final String sheetname = ref.getSheetName();
		final int row = ref.getRow();
		final int col = ref.getCol();
		
		// シート選択してセル選択してContextを返す。
		return this.sheet( sheetname ).cell( row, col );
	}
	


	public RangeContext range(int r1, int c1, int r2, int c2) {
		
		// TODO：一応判定式入れておいた方が良いか。
		
		return new RangeContext( this.current,
				new CellReference( this.current.getSheetName(), r1, c1, false, false ),
				new CellReference( this.current.getSheetName(), r2, c2, false, false ) );
	}
	
	public RangeContext range(String name) {
		
		var xname = this.book.getName( name );
		var ref = new AreaReference( xname.getRefersToFormula(), SpreadsheetVersion.EXCEL2007 );
		
		final String sheetname = xname.getSheetName();
		this.sheet( sheetname ); // シート移動
		
		return new RangeContext( this.current,
				ref.getFirstCell(),
				ref.getLastCell() );
	}
	
	
	@SuppressWarnings("unused")
	@Deprecated
	private XSSFCell xssfCell(CellReference ref) {

		// TODO：厳密指定するオプションを追加したほうが良いか？
		int row = ref.isRowAbsolute() ? ref.getRow() : 0;
		int col = ref.isColAbsolute() ? ref.getCol() : 0;
		
		return this.xssfCell( row, col );
	}
	
	private XSSFCell xssfCell(final int row, final int col) {
		
		XSSFRow xrow = this.current.getRow( row );
		if ( null == xrow ) xrow = this.current.createRow( row );
		
		XSSFCell xcell = xrow.getCell( col );
		if ( null == xcell ) xcell = xrow.createCell( col );
		
		return xcell;
	}
	
	
	public ExcelWrapper shiftRows( final int baseRow, final int shiftSize ) {
		poiShiftRows( this.current, baseRow, shiftSize );
		return this;
	}
	
	public ExcelWrapper hideRows( final int baseRow, final int hideSize ) {
		return this.hideRows( baseRow, hideSize, false );
	}
	public ExcelWrapper hideRows( final int baseRow, final int hideSize, final boolean withClear ) {
		poiHideRows( this.current, baseRow, hideSize, withClear );
		return this;
	}
	
	
	
	public ExcelWrapper copyRows( final int rowSrcTop, final int rowSrcBottom ) {
		return this.copyRows( rowSrcTop, rowSrcBottom, rowSrcBottom + 1 );
	}
	public ExcelWrapper copyRows( final int rowSrcTop, final int rowSrcBottom, final int rowDst ) {
		return this.copyRows( rowSrcTop, rowSrcBottom, rowDst, 1 );
	}
	public ExcelWrapper copyRows( final int rowSrcTop, final int rowSrcBottom, final int rowDst, final int count ) {
		return this.copyRows( rowSrcTop, rowSrcBottom, rowDst, count, DEFAULT_COPY_POLICY );
	}
	public ExcelWrapper copyRows( final int rowSrcTop, final int rowSrcBottom, final int rowDst, final int count, final CellCopyPolicy policy ) {

		final int rows = rowSrcBottom - rowSrcTop + 1; // 植木算
		
		
		// ■シフト処理：
		poiShiftRows( this.current, rowDst, rows * count );
		
		// ■コピー処理：
		for ( int n = 0; n < count; n++ ) {
			
			final int shift = n * rows;
			final int dest = rowDst + shift;
			
			this.current.copyRows( rowSrcTop, rowSrcBottom, dest, policy );
		}
		
		return this;
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
