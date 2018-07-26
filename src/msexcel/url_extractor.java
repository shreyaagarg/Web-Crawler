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

public class url_extractor {
	public static String url;
	static int line = 0;

	public static void main(String[] args) {

		int rownum = 0;
		XSSFWorkbook workbook1 = new XSSFWorkbook();
		XSSFSheet sheet2 = workbook1.createSheet("URL's");
		// html_string.main(args);
		url = html_string.getline();
		Pattern p = Pattern
				.compile("\\b((?:https?|ftp|file)://[-a-zA-Z0-9+&@#/%?=~_|!:,.;]*[-a-zA-Z0-9+&@#/%=~_|])");
		Matcher m = p.matcher(url);
		String temp = null;
		while (m.find()) {
			temp = m.group(1);
			Map<String, Object[]> data = new TreeMap<String, Object[]>();

			data.put("line", new Object[] { temp });
			Set<String> keyset = data.keySet();
			for (String key : keyset) {
				// this creates a new row in the sheet
				Row row = sheet2.createRow(rownum++);
				Object[] objArr = data.get(key);
				int cellnum = 0;
				for (Object obj : objArr) {
					// this line creates a cell in the next column of that row

					Cell cell = row.createCell(cellnum++);
					if (obj instanceof String)
						cell.setCellValue((String) obj);
					else if (obj instanceof Integer)
						cell.setCellValue((Integer) obj);

				}

			}
			System.out.println(temp);
			line++;// this variable should contain the link URL

		}

		try {
			// this Writes the workbook url
			FileOutputStream out = new FileOutputStream(new File("url.xlsx"));
			workbook1.write(out);
			out.close();
			System.out.println("url.xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			more_url.main(args);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	static int get_line() {
		return line;
	}
}
