package excel.out.test;

import org.apache.poi.ss.SpreadsheetVersion;

import co.system.excel.write.WorkbookWriter;

public class ExcelOutputTest {


	public static void main(String[] args) {
		System.out.println("----< obj作成 >------------------");
		try {
			WorkbookWriter work = new WorkbookWriter("C:\\Users\\kazum\\", "Test.xlsx", SpreadsheetVersion.EXCEL2007);

			System.out.println("----< worksheet作成 >------------------");
			String sheet1 = "テストシート１";
			String sheet2 = "テストシート２";
			work.createWorkbookSheet(sheet1);
			work.createWorkbookSheet(sheet2);

			System.out.println("----< cell作成 >------------------");
			String[] eng = { "A", "C", "DF", "CE", "F", "s", "k" };
			int[] nums = { 1, 7, 91, 35, 45, 75, 65, 42, 11 };
			for (String str : eng) {
				for (int num : nums) {
					String out = str + num;
					work.writeCell(sheet1, out, out + "に出力");
				}
				work.setAllColumnSize(sheet1, str, 98);
			}
			work.excelOutput();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
