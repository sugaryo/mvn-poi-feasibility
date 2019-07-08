package sugaryo.poi_feasibility.utility;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiUtil {
	
	public static final byte[] serialize( XSSFWorkbook book ) {
		
		try ( ByteArrayOutputStream stream = new ByteArrayOutputStream() ) {
			
			book.write( stream );
			
			byte[] bin = stream.toByteArray();
			return bin;
		}
		// 検査例外は RuntimeException でラップしてスローする。
		catch ( Exception ex ) {
			throw new RuntimeException( ex );
		}
	}
	
	public static final void output( XSSFWorkbook book, String path ) {
		output( book, Paths.get( path ) );
	}
	
	public static final void output( XSSFWorkbook book, Path path ) {
		
		try {
			Files.createDirectories( path.getParent() );
		}
		// 検査例外は RuntimeException でラップしてスローする。
		catch ( Exception ex ) {
			throw new RuntimeException( ex );
		}
		
		try ( FileOutputStream stream = new FileOutputStream( path.toFile() ) ) {
			book.write( stream );
		}
		// 検査例外は RuntimeException でラップしてスローする。
		catch ( Exception ex ) {
			throw new RuntimeException( ex );
		}
	}
	
	
	/**
	 * Excelテンプレートのバイナリオープン.
	 * 
	 * @param path
	 *            テンプレートファイルのパス
	 * @return 指定したパスのテンプレートを読み込んで {@link XSSFWorkbook} を返す。<br>
	 *         {@code POI version 3.11} からの仕様変更でテンプレートファイルへの直接書き込みになってしまったため、<br>
	 *         ワークブックとテンプレートファイルのFile的な結び付きがないよう、バイナリストリームを介して開く。<br>
	 *         詳細は参考情報の通り。         
	 * 
	 * @see <a href="https://so-kai-app.sakura.ne.jp/blog/1132/2015/12/10/">参考情報 - [POI] 対処編：テンプレートファイルが書き変わるようになってしまった</a>
	 * @see <a href="https://mvnrepository.com/artifact/org.apache.poi/poi">apache poi - Mavenリポジトリ</a>
	 * @see <a href="https://poi.apache.org/changes.html">apache poi - 変更履歴</a>
	 */
	public static final XSSFWorkbook open( String path ) {
		
		final byte[] bin = read( path );
		try ( var stream = new ByteArrayInputStream( bin ) ) {
			return new XSSFWorkbook( stream );
		}
		// 検査例外は RuntimeException でラップしてスローする。
		catch ( Exception ex ) {
			throw new RuntimeException( ex );
		}
	}
	
	private static final byte[] read( String path ) {
		
		try ( var stream = new FileInputStream( new File( path ) ) ) {
			return stream.readAllBytes();
		}
		// 検査例外は RuntimeException でラップしてスローする。
		catch ( Exception ex ) {
			throw new RuntimeException( ex );
		}
	}
	
	/**
	 * @see org.apache.poi.xssf.usermodel.XSSFSheet
	 * @see org.apache.poi.xssf.usermodel.XSSFSheet#shiftRows(int, int, int, boolean, boolean)
	 */
	public static class ShiftRowsOptions {
		
		/**
		 * {@link XSSFSheet#shiftRows(int, int, int, boolean, boolean)}<b>::copyRowHeight.</b> {@code whether to copy the row height during the shift.}
		 * 
		 * @see org.apache.poi.xssf.usermodel.XSSFSheet#shiftRows(int, int, int, boolean, boolean)
		 */
		public static final boolean COPY_ROW_HEIGHT = true;
		/**
		 * {@link XSSFSheet#shiftRows(int, int, int, boolean, boolean)}<b>::resetOriginalRowHeight.</b> {@code whether to set the original row's height to the default.}
		 * 
		 * @see org.apache.poi.xssf.usermodel.XSSFSheet#shiftRows(int, int, int, boolean, boolean)
		 */
		public static final boolean RESET_ORG_ROW_HEIGHT = false; // そもそも poi の中のロジックで使ってないくさいが？
	}
	/**
	 * 行シフト処理.
	 * 
	 * @param sheet     操作する {@link XSSFSheet} オブジェクト
	 * @param baseRow   シフトの基準行インデックス
	 * @param shiftSize シフトする行数
	 * 
	 * @see ShiftRowsOptions
	 * @see ShiftRowsOptions#COPY_ROW_HEIGHT
	 * @see ShiftRowsOptions#RESET_ORG_ROW_HEIGHT
	 */
	public static void poiShiftRows(final XSSFSheet sheet, final int baseRow, final int shiftSize) {
		
		// シフト量が指定されてない場合は無視。
		if ( 0 == shiftSize ) return;
		
		sheet.shiftRows( baseRow, sheet.getLastRowNum(), shiftSize,
				ShiftRowsOptions.COPY_ROW_HEIGHT,
				ShiftRowsOptions.RESET_ORG_ROW_HEIGHT );
	}
}
