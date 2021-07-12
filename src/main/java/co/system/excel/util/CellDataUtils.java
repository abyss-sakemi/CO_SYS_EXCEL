package co.system.excel.util;

import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * セルの情報を設定するクラス
 * @author Abyss-sakemi
 */
public class CellDataUtils {

	/**
	 * A1方式をR1C1方式に直してMapに記載
	 * @param A1方式
	 * @return HashMap key(column) = 行, key(row) = 列 
	 */
	public static HashMap<String, Integer> convertCell(String cell) {
		//正規表現にて行列を区別(小文字の大文字変換)
		String upCell = cell.toUpperCase();
		Pattern p = Pattern.compile("([A-Z]+)([0-9]+)");
		Matcher m = p.matcher(upCell);

		//一致する場合のみ行列で格納
		HashMap<String, Integer> result = new HashMap<String, Integer>();
		if (m.find()) {
			result.put("column", convertColumn(m.group(1)) - 1);
			result.put("row", Integer.parseInt(m.group(2)) - 1);
		}
		return result;
	}

	/**
	 * 列の英語を数字に変換
	 * @param String 列 : 英語
	 * @return int 列 : 数値
	 */
	public static int convertColumn(String value) {
		String upColumn = value.toUpperCase();
		//1文字ずつ取得
		String[] strArray = upColumn.split("");

		//セルの順番を入れ替える。
		List<String> list = Arrays.asList(strArray);
		Collections.reverse(list);
		int result = 0;

		//英語を数字に変換
		for (int i = 0; list.size() > i; i++) {
			int size = StringUtil.convertChar(strArray[i]);
			for (int j = i; j > 0; j--) {
				size *= 26;
			}
			result += size;
		}
		return result;
	}

	/**
	 * 行取得
	 * @param sheet シート
	 * @param rowNum 行番号
	 * @return 行情報
	 */
	public static Row getRow(Sheet sheet, int rowNum) {
		Row row = null;
		try {
			row = sheet.getRow(rowNum);
		} catch (Exception e) {
			//何も行わない。(Rowが使用されていない場合例外発生するため)
		}
		//取得できない場合行を作成する必要がある
		if (row == null) {
			row = sheet.createRow(rowNum);
		}
		return row;
	}

	/**
	 * セル取得
	 * @param row 行データ
	 * @param column 列番号
	 * @return セル情報
	 */
	public static Cell getCell(Row row, int column) {
		Cell cel = null;
		try {
			cel = row.getCell(column);
		} catch (Exception e) {
			//何も行わない。(Rowが使用されていない場合例外発生するため)
		}
		//取得できない場合セルを作成する必要がある
		if (cel == null) {
			cel = row.createCell(column);
		}
		return cel;
	}

	/**
	 * 行のサイズ設定
	 * @param size 行サイズ
	 * @return 行サイズ
	 */
	public static short getRowSize(int rowSize) {
		return (short) (rowSize * 15);
	}

	/**
	 * 列のサイズ設定
	 * @param columnSize 列サイズ
	 * @return 列サイズ
	 */
	public static int getColumnSize(int columnSize) {
		return columnSize * 32;
	}
}
