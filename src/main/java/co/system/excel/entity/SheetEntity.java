package co.system.excel.entity;

import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Sheet;

import lombok.Data;

/**
 * シート情報格納クラス
 * @author abyss
 */
@Data
public class SheetEntity {

	/** シート番号 */
	private int sheetNum;
	/** シート名 */
	private String sheetName;
	/** セル情報 */
	private Map<String, String> cellsMap;
	/** シート情報 */
	private Sheet sheet;

	/**
	 * コンストラクタ
	 * @param sheetNum シート番号
	 * @param sheetName シート名
	 * @param sheet シート情報
	 */
	public SheetEntity(int sheetNum, String sheetName, Sheet sheet) {
		this.sheetNum = sheetNum;
		this.sheetName = sheetName;
		this.sheet = sheet;
		this.cellsMap = new TreeMap<String, String>();
	}
}
