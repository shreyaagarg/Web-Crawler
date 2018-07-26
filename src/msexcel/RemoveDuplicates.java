package msexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.StringTokenizer;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * TODO To change the template for this generated type comment go to Window -
 * Preferences - Java - Code Style - Code Templates
 */
public class RemoveDuplicates {

	/**
  *  
  */
	public RemoveDuplicates() {
		super();
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) {
		XSSFWorkbook workBook = null;
		FileInputStream fs = null;
		XSSFSheet sheet = null;
		// File file = new File(".");
		// for(String fileNames : file.list()) System.out.println(fileNames);
		try {
			fs = new FileInputStream(new File("new url.xlsx"));
			workBook = new XSSFWorkbook(fs);
			sheet = workBook.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			System.out.println(rows);
			Set s = new HashSet();
			String str = "";
			for (int i = 0; i < rows; i++) {
				str = "";
				XSSFRow row = sheet.getRow((short) i);
				int columns = row.getPhysicalNumberOfCells();
				for (int j = 0; j < columns; ++j) {
					XSSFCell cell0 = row.getCell((short) j);
					int type = cell0.getCellType();
					if (type == 0) {
						double intValue = cell0.getNumericCellValue();
						str = str + String.valueOf(intValue) + ",";
					} else if (type == 1) {
						String stringValue = cell0.getStringCellValue();
						str = str + stringValue + ",";
					}

				}
				str = str.replace(str.charAt(str.lastIndexOf(",")), ' ');
				s.add(str.trim());
			}
			StringTokenizer st = null;
			String result = "";
			Iterator iter = s.iterator();

			// Create a new workbook for the output excel
			XSSFWorkbook workBookOut = new XSSFWorkbook();

			// Create a new Sheet in the output excel workbook
			XSSFSheet sheetOut = workBookOut.createSheet("Remove Duplicates");
			XSSFRow[] row = new XSSFRow[s.size()];
			int rowCount = 0;
			int cellCount = 0;
			while (iter.hasNext()) {
				cellCount = 0;
				row[rowCount] = sheetOut.createRow(rowCount);
				result = iter.next().toString();
				System.out.println(result);
				st = new StringTokenizer(result, " ");
				XSSFCell[] cell = new XSSFCell[st.countTokens()];
				while (st.hasMoreTokens()) {
					cell[cellCount] = row[rowCount]
							.createCell((short) cellCount);
					cell[cellCount].setCellValue(st.nextToken());
					++cellCount;
				}
				++rowCount;
			}
			FileOutputStream fileOut = new FileOutputStream(new File(
					"final.xlsx"));
			workBookOut.write(fileOut);
			fileOut.close();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		}
		System.out.println("final.xlsx written successfully on disk.");
	}
}