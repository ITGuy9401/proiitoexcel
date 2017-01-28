package eu.arcangelovicedomini.proiiexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;

public class Application {

	public static void main(String[] args) throws Exception {
		System.out
				.println("Pro II to Excel Converter.\nI'll pick the file named \"proii.plt\" in the folder of this app "
						+ "and I'll create a proii.xlsx file in the same folder");
		
		File pltFile = new File("proii.plt");
		if (!pltFile.exists()) {
			new FileNotFoundException("I wasn't able to find the file \"proii.plt\"").printStackTrace();
		}
		
		FileInputStream fis = new FileInputStream(pltFile);
		
		List<String> plt = IOUtils.readLines(fis);
		
		List<Point> bubblePoints = new ArrayList<Point>();
		List<Point> dewPoints = new ArrayList<Point>();
		
		boolean reachedDewPoints = false;
		boolean finishedReading = false;
		for (int i = 0; i < plt.size() && !finishedReading; i++) {
			String line = plt.get(i);
			if(!reachedDewPoints) {
				if (line.indexOf("CURVE") != -1) {
					line = plt.get(++i);
					System.out.println("Found " + stripAndTrim(line) + " Bubble Points");
					line = stripAndTrim(plt.get(++i));
					do {
						String[] split = line.split(" ");
						bubblePoints.add(new Point(Double.parseDouble(split[0]), Double.parseDouble(split[1])));
						line = stripAndTrim(plt.get(++i));
					} while(line.indexOf("MARKER") != -1);
					reachedDewPoints = true;
				}
			} else {
				if (line.indexOf("CURVE") != -1) {
					line = plt.get(++i);
					System.out.println("Found " + stripAndTrim(line) + " Bubble Points");
					line = stripAndTrim(plt.get(++i));
					do {
						String[] split = line.split(" ");
						dewPoints.add(new Point(Double.parseDouble(split[0]), Double.parseDouble(split[1])));
						line = stripAndTrim(plt.get(++i));
					} while(line.indexOf("MARKER") != -1);
					finishedReading = true;
				}
			}
		}
		
		
	}

	private static String stripAndTrim(String line) {
		return StringUtils.strip(StringUtils.trim(line));
	}
}
