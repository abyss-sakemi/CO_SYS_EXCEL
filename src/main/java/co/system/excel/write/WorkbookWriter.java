package co.system.excel.write;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import co.system.excel.entity.SheetEntity;
import co.system.excel.util.CellDataUtils;

/**
 * Excel出力Util
 * @author abyss-sakemi
 */
public class WorkbookWriter {
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

	/**
	 * テンプレートファイル読込
	 * @param inPath ファイルパス
	 * @param inName ファイル名
	 * @throws Exception 例外
	 */
	public Workbook WorkbookWriterMake(String inPath, String inName) throws Exception {
		return template(inPath, inName);
	}

	/**
	 * 引数付きコンストラクタ
	 * @param String EXCEL出力先パス
	 * @param String EXCEL出力ファイル名
	 * @param String "HSSF":xls形式 "XSSF":xlsx形式
	 * @throws Exception 例外
	 */
	public WorkbookWriter(String path, String name, SpreadsheetVersion type) throws Exception {
		setOutputFile(path, name);
		createWorkbook(path, name, type);
		sheetMap = new TreeMap<String, SheetEntity>();
	}

	/**
	 * Excelの拡張子ごとのワークブック作成
	 * @param String EXCEL出力先パス
	 * @param String EXCEL出力ファイル名
	 * @param String "HSSF":xls形式 "XSSF":xlsx形式
	 * @throws Exception 例外
	 */
	private void createWorkbook(String getpath, String name, SpreadsheetVersion version) throws Exception {

		System.out.println("**処理開始**ワークブック作成****");

		setOutputFile(path, file);
		String messege;
		switch (version) {
		case EXCEL97:
			workbook = new HSSFWorkbook();
			messege = "HSSFWorkbook(xls形式)";
			break;
		case EXCEL2007:
			workbook = new XSSFWorkbook();
			messege = "XSSFWorkbook(xlsx形式)";
			break;
		default:
			throw new IllegalArgumentException(version.toString());
		}

		System.out.println("**正常終了**[" + messege + "]で作成****");
	}

	/**
	 * ワークシート作成
	 * @param String ワークシート名
	 * @throws Exception 例外
	 */
	public void createWorkbookSheet(String sheetName) throws Exception {
		String trimSheetName = sheetName.trim();
		Sheet sheetInfo;
		System.out.println("**処理開始**[" + file + "] に [" + trimSheetName + "]シートを作成****");

		if (workbook instanceof HSSFWorkbook) {
			sheetInfo = createHSSFWorkbookSheet((HSSFWorkbook) workbook, count);
		} else if (workbook instanceof XSSFWorkbook) {
			sheetInfo = createXSSFWorkbookSheet((XSSFWorkbook) workbook, count);
		} else {
			throw new Exception("Workbookに指定と違う型で作成されています。 xls形式:[HSSF] xlsx形式:[XSSF]");
		}
		SheetEntity sheet = new SheetEntity(count, trimSheetName, sheetInfo);
		sheetMap.put(trimSheetName, sheet);

		System.out.println("**正常終了**[" + file + "] に [" + sheetName.trim() + "]シートを作成(補足 :[" + count + "])****");

		count += 1;

		return;
	}

	/**
	 * 指定したシートのセルに値を書き込む
	 * @param String シート名称
	 * @param String セルの番号(A1方式)
	 * @param String セルに書き込む値
	 * @throws Exception 例外
	 */
	public void writeCell(String sheetName, String cell, String value) throws Exception {

		//cellのA1方式をR1C1方式で取得(rowで行の数値, columnで列の数値)
		HashMap<String, Integer> cells = CellDataUtils.convertCell(cell);

		//セルの取得(シート->列->セルとして取得)
		String trimSheetName = sheetName.trim();
		Sheet nowSheet = sheetMap.get(trimSheetName).getSheet();
		Row row = CellDataUtils.getRow(nowSheet, cells.get("row"));
		Cell cel = CellDataUtils.getCell(row, cells.get("column"));

		//セルに値を設定する。
		cel.setCellValue(value);

		System.out.println("**正常終了**[" + trimSheetName + "] シートの [" + cell + "] : [" + value + "]で作成****");
	}

	/**
	 * 指定シートセルの結合
	 * https://blog.java-reference.com/poi-join-cell/
	 * @param String シート名
	 * @param String 選択セル開始位置
	 * @param String 選択セル最終位置
	 * @throws Exception 例外
	 */
	public void joinCell(String sheetName, String firstCell, String lastCell) throws Exception {
		//結合開始・終了セルの数値を設定
		String trimSheetName = sheetName.trim();
		HashMap<String, Integer> firstCells = CellDataUtils.convertCell(firstCell);
		HashMap<String, Integer> lastCells = CellDataUtils.convertCell(lastCell);

		//開始rowの数値,終了rowの数値,開始columnの数値,終了columnの数値で結合できる。
		int[] firlasCells = { firstCells.get("row"), lastCells.get("row"), firstCells.get("column"),
				lastCells.get("column") };

		System.out.println("**処理開始**[" + trimSheetName + "] シートの [" + firstCell + "]～[" + lastCell + "]で結合****");
		Sheet nowSheet;
		if (workbook instanceof HSSFWorkbook) {
			//xls形式のWorkbookを作成
			nowSheet = (HSSFSheet) sheetMap.get(trimSheetName).getSheet();
		} else if (workbook instanceof XSSFWorkbook) {
			//xlsx形式のWorkbookを作成
			nowSheet = (XSSFSheet) sheetMap.get(trimSheetName).getSheet();
		} else {
			throw new Exception("セルの入力に失敗しました。");
		}
		CellRangeAddress cra = new CellRangeAddress(firlasCells[0], firlasCells[1], firlasCells[2], firlasCells[3]);
		nowSheet.addMergedRegion(cra);
		System.out.println("**正常終了**[" + trimSheetName + "] シートの [" + firstCell + "]～[" + lastCell + "]で結合****");
	}

	/**
	 * シートの指定行のサイズを設定する。
	 * @param String シート名
	 * @param String 変更するrowの英語値
	 * @param float  row(行)のサイズを設定
	 */
	public void setRowSize(String sheetName, String row, int rowSize) throws Exception {

		String trimSheetName = sheetName.trim();
		System.out.println("**処理開始**[" + trimSheetName + "] シートの [" + row + "]のサイズを[" + rowSize + "]に変更****");

		//設定したセルを行う
		setRowSize(trimSheetName, CellDataUtils.convertColumn(row), rowSize);
		System.out.println("**正常終了**[" + trimSheetName + "] シートの [" + row + "]のサイズを[" + rowSize + "]に変更****");

	}

	/**
	 * シートの指定範囲行のサイズを設定する。
	 * @param String シート名
	 * @param int column(列)のサイズを設定
	 */
	public void setRowSize(String sheetName, String firstRow, String lastRow, int rowSize) throws Exception {

		String trimSheetName = sheetName.trim();
		System.out.println(
				"**処理開始**[" + trimSheetName + "] シートの [" + firstRow + "]～[" + lastRow + "]のサイズを[" + rowSize
						+ "]に変更****");

		int firRow = CellDataUtils.convertColumn(firstRow);
		int lasRow = CellDataUtils.convertColumn(lastRow);

		for (int i = firRow; i < lasRow; i++) {
			setRowSize(trimSheetName, i, rowSize);
		}

		System.out.println(
				"**正常終了**[" + trimSheetName + "] シートの [" + firstRow + "]～[" + lastRow + "]のサイズを[" + rowSize
						+ "]に変更****");
	}

	/**
	 * 	シートの指定行のサイズを設定する。
	 * @param String シート名
	 * @param String 変更するrowの数値
	 * @param float  row(行)のサイズを設定
	 */
	public void setRowSize(String sheetName, int rowNum, int rowSize) throws Exception {
		String trimSheetName = sheetName.trim();
		Row row = CellDataUtils.getRow(sheetMap.get(trimSheetName).getSheet(), rowNum--);
		row.setHeight(CellDataUtils.getRowSize(rowSize));
	}

	/**
	 * シートの指定列のサイズを設定する。
	 * @param String シート名
	 * @param int 指定列
	 * @param int column(列)のサイズを設定
	 */
	public void setColumnSize(String sheetName, int column, int columnSize) throws Exception {

		//シートを取得
		String trimSheetName = sheetName.trim();
		Sheet getSheet = sheetMap.get(trimSheetName).getSheet();

		System.out.println("**処理開始**[" + trimSheetName + "] シートの [" + column + "]のサイズを[" + columnSize + "]に変更****");

		//指定列のサイズを設定
		getSheet.setColumnWidth(column, CellDataUtils.getColumnSize(columnSize));

		System.out.println("**正常終了**[" + trimSheetName + "] シートの [" + column + "]のサイズを[" + columnSize + "]に変更****");

	}

	/**
	 * 	指定列のサイズを設定する。
	 * @param String シート名
	 * @param int column(列)番号
	 * @param int column(列)のサイズを設定
	 * @throws Exception 例外
	 */
	public void setAllColumnSize(String sheetName, String column, int columnSize) throws Exception {

		//全列のサイズを設定
		String trimSheetName = sheetName.trim();
		Sheet getSheet = sheetMap.get(trimSheetName).getSheet();
		getSheet.setColumnWidth(CellDataUtils.convertColumn(column) - 1, columnSize * 32);
	}

	/**
	 * テンプレートファイルの読み込み
	 * @param String 入力ファイルパス
	 * @param String 入力ファイル名
	 * @throws Exception 例外
	 */
	public Workbook template(String inPath, String inName) throws Exception {
		try {
			//Excelファイルの読み込み
			workbook = WorkbookFactory.create(new FileInputStream(inPath + inName));
			System.out.println("シートの全シート数 : " + workbook.getNumberOfSheets());
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				//現在のシート情報を取得
				SheetEntity sheet = new SheetEntity(count, workbook.getSheetName(count), workbook.getSheetAt(count));
				sheetMap.put(workbook.getSheetName(count), sheet);
				System.out.println("[" + inName + "]から[" + workbook.getSheetName(count) + "]シートを取得");
				System.out.println("[" + count + "]に[" + workbook.getSheetName(count) + "]シートを格納");
				count += 1;

			}
		} catch (Exception e) {
			throw new Exception("テンプレートの読み込みに失敗しました");
		}
		return workbook;

	}

	/**
	 * Excelの出力ファイル名を設定
	 * @param String 出力ファイルパス
	 * @param String 出力ファイル名
	 */
	public void setOutputFile(String outPath, String outName) {
		path = outPath;
		file = outName;
	}

	/**
	 * Excelの出力を行います。
	 * @param String 出力パス
	 * @param String 出力ファイル名
	 */
	public void excelOutput() {

		//各シートのシート名称を修正
		for (String sheetName : sheetMap.keySet()) {
			workbook.setSheetName(sheetMap.get(sheetName).getSheetNum(), sheetName);
		}

		FileOutputStream outExcelFile = null;
		try {
			// ファイルを出力
			outExcelFile = new FileOutputStream(path + file);
			workbook.write(outExcelFile);
		} catch (Exception e) {
			System.out.println(e.toString());
		} finally {
			try {
				if (outExcelFile != null) {
					outExcelFile.close();
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * xls形式(EXCEL97)のシート作成
	 * @param workbook ワークブック
	 * @param count シート番号
	 * @return シート情報
	 * @throws Exception 例外
	 */
	private static HSSFSheet createHSSFWorkbookSheet(HSSFWorkbook workbook, int count) throws Exception {
		//シートの作成
		HSSFSheet argsSheet = (HSSFSheet) workbook.createSheet();
		argsSheet = (HSSFSheet) workbook.getSheetAt(count);
		return argsSheet;
	}

	/**
	 * xlsx形式(EXCEL2007)のシート作成
	 * @param workbook ワークブック
	 * @param count シート番号
	 * @return シート情報
	 * @throws Exception 例外
	 */
	private static XSSFSheet createXSSFWorkbookSheet(XSSFWorkbook workbook, int count) throws Exception {
		//シートの作成
		XSSFSheet argsSheet = (XSSFSheet) workbook.createSheet();
		argsSheet = (XSSFSheet) workbook.getSheetAt(count);
		return argsSheet;
	}

}