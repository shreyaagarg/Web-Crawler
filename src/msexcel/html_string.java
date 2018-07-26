package msexcel;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;

public class html_string {
	static String line, blank = "";
	public static String hyperlink = "https://www.pec.ac.in";

	public static void main(String[] args) {
		URL url;
		InputStream is = null;
		BufferedReader br;

		try {
			url = new URL(hyperlink);//extract whole source code
			is = url.openStream(); // throws an IOException
			br = new BufferedReader(new InputStreamReader(is));

			while ((line = br.readLine()) != null) {
				blank = blank.concat(" " + line);
				System.out.println(line);
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
		// System.out.println(blank.charAt(0));
	}

	public static String getline() {
		return blank;
	}

}
