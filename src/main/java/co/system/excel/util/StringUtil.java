package co.system.excel.util;

import java.util.HashMap;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 文字列変換共通クラス
 * @author Abyss-sakemi
 */
public class StringUtil {
	
	/** 英語,数字列挙型 */
	@AllArgsConstructor 
	@Getter
	public enum EngChange{
		A("A",1),
		B("B",2),
		C("C",3),
		D("D",4),
		E("E",5),
		F("F",6),
		G("G",7),
		H("H",8),
		I("I",9),
		J("J",10),
		K("K",11),
		L("L",12),
		M("M",13),
		N("N",14),
		O("O",15),
		P("P",16),
		Q("Q",17),
		R("R",18),
		S("S",19),
		T("T",20),
		U("U",21),
		V("V",22),
		W("W",23),
		X("X",24),
		Y("Y",25),
		Z("Z",26),
		;

		/** 英語 */
		private String eng;
		/** 数字 */
		private int num;
	}

	/**  MapValue */
	private static HashMap<String, Integer> map = new HashMap<String, Integer>();

	/**
	 * 初回時格納処理
	 */
	static {
		for (EngChange eng : EngChange.values()) {
			map.put(eng.getEng(), eng.getNum());
		}
	}

	/**
	 * A-Zを1～26で返却
	 * @return 数字に対する1～26の値
	 */
	public static int convertChar(String key) {
		return map.get(key);
	}
}
