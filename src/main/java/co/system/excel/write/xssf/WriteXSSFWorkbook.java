package co.system.excel.write.xssf;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * xlsx形式のWorkbook書込処理(EXCEL2007)
 * @author Abyss-sakemi
 */
public class WriteXSSFWorkbook {

	public static XSSFSheet createXSSFWorkbookSheet(XSSFWorkbook workbook, int count) throws Exception {

		//シートの作成
		XSSFSheet argsSheet = (XSSFSheet) workbook.createSheet();
		argsSheet = (XSSFSheet) workbook.getSheetAt(count);

		return argsSheet;
	}
}