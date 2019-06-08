package sugaryo.poi_feasibility.utility;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.FileAttribute;

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
}
