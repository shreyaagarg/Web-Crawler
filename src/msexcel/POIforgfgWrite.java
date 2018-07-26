package msexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIforgfgWrite {
	public static String code;
	public static String url;

	public static void main(String[] args) {

		html_string.main(args);

		code = html_string.getline();// get the whole html code in string
		url = html_string.getline();
		// Blank workbook

		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create a blank sheet
		XSSFSheet sheet1 = workbook.createSheet("HTML Code");
		// XSSFSheet sheet2 = workbook.createSheet("URL's");
		// This data needs to be written (Object[])
		int rownum = 0;
		int line = 1;
		String temp = " ";
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		Set<String> keyset = data.keySet();
		for (int i = 0; i < code.length(); i++) {
			temp = "";
			// inserting tags in to excel file
			if (code.charAt(i) == '<') {
				// System.out.println("test");
				i++;
				while (code.charAt(i) != '>') {
					temp = temp + Character.toString(code.charAt(i));//separate each tag into a string
					i++;
					// System.out.println(temp );
				}
				// i++;
				System.out.println(temp);

				data.put("line", new Object[] { temp });

				for (String key : keyset) {
					// this creates a new row in the sheet
					Row row = sheet1.createRow(rownum++);
					Object[] objArr = data.get(key);
					int cellnum = 0;
					for (Object obj : objArr) {
						// sthis line creates a cell in the next column of that
						// row

						Cell cell = row.createCell(cellnum++);
						if (obj instanceof String)
							cell.setCellValue((String) obj);
						else if (obj instanceof Integer)
							cell.setCellValue((Integer) obj);
					}

					/*
					 * Map<String, Object[]> data = new TreeMap<String,
					 * Object[]>();
					 * 
					 * data.put("line", new Object[]{ temp }); line++;
					 * 
					 * // Iterate over data and write to sheet Set<String>
					 * keyset = data.keySet();
					 * 
					 * for (String key : keyset) { // this creates a new row in
					 * the sheet Row row = sheet1.createRow(rownum++); Object[]
					 * objArr = data.get(key); int cellnum = 0; for (Object obj
					 * : objArr) { // this line creates a cell in the next
					 * column of that row
					 * 
					 * Cell cell = row.createCell(cellnum++); if (obj instanceof
					 * String) cell.setCellValue((String)obj); else if (obj
					 * instanceof Integer) cell.setCellValue((Integer)obj); }
					 */
				}

				/*
				 * try { // this Writes the workbook gfgcontribute
				 * FileOutputStream out = new FileOutputStream(new
				 * File("gfgcontribute.xlsx")); workbook.write(out);
				 * out.close(); System.out.println(
				 * "gfgcontribute.xls written successfully on disk." ); } catch
				 * (Exception e) { e.printStackTrace(); }
				 */
				line++;
				continue;
			}

			else if (i != 0 && code.charAt(i - 1) == '>') {
				// System.out.println("test");
				// i++;
				while (code.charAt(i + 1) != '<') {
					temp = temp + Character.toString(code.charAt(i));
					i++;
					// System.out.println(temp );
				}
				temp = temp + Character.toString(code.charAt(i));
				System.out.println(temp);
				// Map<String, Object[]> data = new TreeMap<String, Object[]>();

				data.put("line", new Object[] { temp });
				// Set<String> keyset = data.keySet();
				for (String key : keyset) {
					// this creates a new row in the sheet
					Row row = sheet1.createRow(rownum++);
					Object[] objArr = data.get(key);
					int cellnum = 0;
					for (Object obj : objArr) {
						// this line creates a cell in the next column of that
						// row

						Cell cell = row.createCell(cellnum++);
						if (obj instanceof String)
							cell.setCellValue((String) obj);
						else if (obj instanceof Integer)
							cell.setCellValue((Integer) obj);
					}

					/*
					 * Map<String, Object[]> data = new TreeMap<String,
					 * Object[]>();
					 * 
					 * data.put("line", new Object[]{ temp }); line++;
					 * 
					 * // Iterate over data and write to sheet Set<String>
					 * keyset = data.keySet();
					 * 
					 * for (String key : keyset) { // this creates a new row in
					 * the sheet Row row = sheet1.createRow(rownum++); Object[]
					 * objArr = data.get(key); int cellnum = 0; for (Object obj
					 * : objArr) { // this line creates a cell in the next
					 * column of that row
					 * 
					 * Cell cell = row.createCell(cellnum++); if (obj instanceof
					 * String) cell.setCellValue((String)obj); else if (obj
					 * instanceof Integer) cell.setCellValue((Integer)obj); }
					 */
				}

				/*
				 * try { // this Writes the workbook gfgcontribute
				 * FileOutputStream out = new FileOutputStream(new
				 * File("gfgcontribute.xlsx")); workbook.write(out);
				 * out.close(); System.out.println(
				 * "gfgcontribute.xls written successfully on disk." ); } catch
				 * (Exception e) { e.printStackTrace(); }
				 */
				line++;
				continue;
			}
		}
		try {
			// this Writes the workbook gfgcontribute
			FileOutputStream out = new FileOutputStream(new File(
					"gfgcontribute.xlsx"));
			workbook.write(out);
			out.close();
			System.out
					.println("gfgcontribute.xls written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
		/*
		 * Pattern p = Pattern.compile("href=\"(.*?)\""); Matcher m =
		 * p.matcher(url); temp = null; line=0; while (m.find()) { temp =
		 * m.group(1); // Map<String, Object[]> data = new TreeMap<String,
		 * Object[]>();
		 * 
		 * 
		 * data.put("line", new Object[]{ temp }); Set<String> keyset =
		 * data.keySet(); for (String key : keyset) { // this creates a new row
		 * in the sheet Row row = sheet2.createRow(rownum++); Object[] objArr =
		 * data.get(key); int cellnum = 0; for (Object obj : objArr) { // this
		 * line creates a cell in the next column of that row
		 * 
		 * Cell cell = row.createCell(cellnum++); if (obj instanceof String)
		 * cell.setCellValue((String)obj); else if (obj instanceof Integer)
		 * cell.setCellValue((Integer)obj); try { // this Writes the workbook
		 * gfgcontribute FileOutputStream out = new FileOutputStream(new
		 * File("gfgcontribute.xlsx")); workbook.write(out); // out.close();
		 * System.out.println("gfgcontribute.xlsx written successfully on disk."
		 * ); } catch (Exception e) { e.printStackTrace(); } }
		 * 
		 * }
		 * 
		 * System.out.println(temp ); line++; }
		 */
		/*
		 * try { more_url.main(args); } catch (IOException e) { // TODO
		 * Auto-generated catch block e.printStackTrace(); }
		 */
		url_extractor.main(args);
	}
}
