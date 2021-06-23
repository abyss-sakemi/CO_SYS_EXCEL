package co.system.excel.write.hssf;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * xls形式のWorkbook書込処理(EXCEL97)
 * @author Abyss
 */
public class WriteHSSFWorkbook {

	public static HSSFSheet createHSSFWorkbookSheet(HSSFWorkbook workbook, int count) throws Exception {

		//シートの作成
		HSSFSheet argsSheet = (HSSFSheet) workbook.createSheet();
		argsSheet = (HSSFSheet) workbook.getSheetAt(count);

		return argsSheet;
	}
}