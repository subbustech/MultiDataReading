package datareading;

import java.util.Map;

public class ReadEmpDetails {

	public static void main(String[] args) {

		//GetDataFromExcelWithKey();
		GetDataFromExcelWithoutKey();

	}
	public static void GetDataFromExcelWithKey() {
		String path = "data/differentdata.xlsx";
		String sheetName = "Sheet1";
		String key = "empdetails";
		MultiDataReading mdr = new MultiDataReading();
		Map<Integer, Map<String, String>> data = mdr.getData(path, sheetName, key);

		System.out.println(data);

		for(int i=1; i<=data.size(); i++) {
			//System.out.println(data.get(i));
			System.out.println(data.get(i).get("id"));
			System.out.println(data.get(i).get("firstname"));
			System.out.println(data.get(i).get("lastname"));
			System.out.println(data.get(i).get("company"));
			System.out.println(data.get(i).get("role"));
			System.out.println(data.get(i).get("salary"));

			boolean success = mdr.setData(path, sheetName, key, i, "company", "12345");
			if(success) {
				System.out.println("Sheet updated successfully");
			}
		}	
	}
	public static void GetDataFromExcelWithoutKey() {
		String path = "data/differentdata1.xlsx";
		String sheetName = "Sheet1";
		MultiDataReading mdr = new MultiDataReading();
		Map<Integer, Map<String, String>> data = mdr.getData(path, sheetName);

		System.out.println(data);

		for(int i=1; i<=data.size(); i++) {
			//System.out.println(data.get(i));
			System.out.println(data.get(i).get("id"));
			System.out.println(data.get(i).get("firstname"));
			System.out.println(data.get(i).get("lastname"));
			System.out.println(data.get(i).get("company"));
			System.out.println(data.get(i).get("role"));
			System.out.println(data.get(i).get("salary"));
			
			boolean success = mdr.setData(path, sheetName, i, "password", "test123");
			if(success) {
				System.out.println("Sheet updated successfully");
			}
		}
	}
}
