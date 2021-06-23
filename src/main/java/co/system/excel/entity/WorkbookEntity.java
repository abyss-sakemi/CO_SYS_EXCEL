package co.system.excel.entity;

import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Workbook;

import lombok.Data;

@Data
/**
 * ワークブック情報格納クラス
 * @author Abyss-sakemi
 */
public class WorkbookEntity {

	/** パス */
	private String path;

	/** ファイル名 */
	private String file;

	/** ワークブック */
	private Workbook workbook;

	/** シートカウント */
	private int count = 0;

	/** シートマップ */
	private TreeMap<String, SheetEntity> sheetMap;

}
