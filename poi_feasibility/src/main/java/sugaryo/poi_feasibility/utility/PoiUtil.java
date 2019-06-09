package sugaryo.poi_feasibility.utility;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

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
			return (XSSFWorkbook)XSSFWorkbookFactory.create( stream );
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
}
