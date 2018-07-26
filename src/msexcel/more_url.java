package msexcel;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;

import msexcel.html_string;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class more_url {
	static String liner, blank = "";
	static int line = url_extractor.get_line();

	public static void main(String[] args) throws IOException {
		// url_extractor.main(args);
		try {
			FileInputStream file = new FileInputStream(new File("url.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						cell.getNumericCellValue();
						break;
					case Cell.CELL_TYPE_STRING: {
						URL url;
						InputStream is = null;
						BufferedReader br;

						try {
							url = new URL(cell.getStringCellValue());
							is = url.openStream(); // throws an IOException
							br = new BufferedReader(new InputStreamReader(is));

							while ((liner = br.readLine()) != null) {
								blank = blank.concat(" " + liner);
								// System.out.println(liner);
							}
						} catch (MalformedURLException mue) {
							mue.printStackTrace();
						} catch (IOException ioe) {
							ioe.printStackTrace();
						} finally {
							try {
								if (is != null)
									is.close();
							} catch (IOException ioe) {
								// nothing to see here
							}
						}
						int rownum = 0;
						XSSFWorkbook workbook1 = new XSSFWorkbook();
						XSSFSheet sheet2 = workbook1.createSheet("URL's");
						html_string.main(args);
						//calling the url
						HashMap<String, Object[]> data1 = new HashMap<String, Object[]>();

						data1.put("line",
								new Object[] { cell.getStringCellValue() });
						Set<String> keyset1 = data1.keySet();
						for (String key : keyset1) {
							// this creates a new row in the sheet
							Row row1 = sheet2.createRow(rownum++);
							Object[] objArr = data1.get(key);
							int cellnum = 0;
							for (Object obj : objArr) {
								// this line creates a cell in the next column
								// of that row

								Cell cell1 = row1.createCell(cellnum++);
								if (obj instanceof String)
									cell1.setCellValue((String) obj);
								else if (obj instanceof Integer)
									cell1.setCellValue((Integer) obj);

							}

						}
						// System.out.println(temp );
						line++;// this variable should contain the link URL
						// url=html_string.getline();
						Pattern p = Pattern
								.compile("\\b((?:https?|ftp|file)://[-a-zA-Z0-9+&@#/%?=~_|!:,.;]*[-a-zA-Z0-9+&@#/%=~_|])");
						Matcher m = p.matcher(blank);
						String temp = null;
						HashMap<String, Object[]> data = new HashMap<String, Object[]>();

						while (m.find()) {
							temp = m.group(1);

							// if(data.containsValue(temp)==false){
							data.put("line", new Object[] { temp });
							Set<String> keyset = data.keySet();
							for (String key : keyset) {
								// this creates a new row in the sheet
								Row row1 = sheet2.createRow(rownum++);
								Object[] objArr = data.get(key);
								int cellnum = 0;
								for (Object obj : objArr) {
									// this line creates a cell in the next
									// column of that row

									Cell cell1 = row1.createCell(cellnum++);
									if (obj instanceof String)
										cell1.setCellValue((String) obj);
									else if (obj instanceof Integer)
										cell1.setCellValue((Integer) obj);

								}

							}
							// }
							System.out.println(temp + " "
									+ data.containsValue(temp));
							line++;// this variable should contain the link URL

						}

						try {
							// this Writes the workbook url
							FileOutputStream out = new FileOutputStream(
									new File("new url.xlsx"));
							workbook1.write(out);
							out.close();
							System.out
									.println("new url.xlsx written successfully on disk.");
						} catch (Exception e) {
							e.printStackTrace();
						}
					}

						break;

					}

					// System.out.println(blank);
				}

				file.close();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		RemoveDuplicates.main(args);
	}

}
