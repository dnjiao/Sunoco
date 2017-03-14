package accounting.gulfmark.services;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import accounting.gulfmark.models.VolEntry;


public class VolumeParser {
	public static List<VolEntry> getEntries(String text) {
		List<VolEntry> entries = new ArrayList<VolEntry>();
		String[] lines = text.split("\\r?\\n");
		for(int i = 0; i < lines.length; i++) {
			String line = lines[i];
			if (lines[i].startsWith("***")) {
				VolEntry volEntry = new VolEntry();
				int firstSpace = line.indexOf(" ");
				int secondSpace = line.indexOf(" ", firstSpace + 1);
				volEntry.setVolNumber(line.substring(firstSpace + 1, secondSpace));
				int idxCounty = line.indexOf("COUNTY");
				volEntry.setVolName(line.substring(secondSpace + 1, idxCounty));
				int firstColon = line.indexOf(":", idxCounty);
				int endCounty = line.indexOf(" ", firstColon);
				volEntry.setCounty(line.substring(firstColon + 1, endCounty));
				int secondColon = line.indexOf(":", endCounty);
				volEntry.setState(line.substring(secondColon + 1).trim());
				i++;
				if (lines[i].startsWith("W ")) {
					String[] items = lines[i].trim().split(" +");
					volEntry.setMonth(items[2]);
					volEntry.setYear(items[3]);
					double net;
					if (items.length == 6) {
						if (items[5].endsWith("-")) {
							net = Double.parseDouble(items[5].substring(0, items[5].indexOf('-'))) * (-1);
						} else {
							net = Double.parseDouble(items[5]);
						}
						volEntry.setNet(net);
					} else if (items.length == 8){
						volEntry.setUnit(Double.parseDouble(items[4]));
						volEntry.setVolume(Double.parseDouble(items[5]));
						volEntry.setNet(Double.parseDouble(items[6]));
					}
				}
				entries.add(volEntry);	
			}
		}
		return entries;
	}
	

	
	public static void createXls(List<VolEntry> entries, String xlsPath) throws Exception {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Volume");
		
		// Create title row
		int rowId = 0;
		HSSFRow row = sheet.createRow(rowId++);
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("Month");
		cell = row.createCell(1);
		cell.setCellValue("Year");
		cell = row.createCell(2);
		cell.setCellValue("Volume #");
		cell = row.createCell(3);
		cell.setCellValue("Volume Name");
		cell = row.createCell(4);
		cell.setCellValue("County");
		cell = row.createCell(5);
		cell.setCellValue("State");
		cell = row.createCell(6);
		cell.setCellValue("Unit");
		cell = row.createCell(7);
		cell.setCellValue("Volume");
		cell = row.createCell(8);
		cell.setCellValue("Net");
		for (VolEntry entry : entries) {
			row = sheet.createRow(rowId++);
			cell = row.createCell(0);
			cell.setCellValue(entry.getMonth());
			cell = row.createCell(1);
			cell.setCellValue(entry.getYear());
			cell = row.createCell(2);
			cell.setCellValue(entry.getVolNumber());
			cell = row.createCell(3);
			cell.setCellValue(entry.getVolName());
			cell = row.createCell(4);
			cell.setCellValue(entry.getState());
			cell = row.createCell(5);
			cell.setCellValue(entry.getCounty());
			if (entry.getUnit() != null) {
				cell = row.createCell(6);
				cell.setCellValue(entry.getUnit());
			}
			if (entry.getVolume() != null) {
				cell = row.createCell(7);
				cell.setCellValue(entry.getVolume());
			} 
			cell = row.createCell(8);
			cell.setCellValue(entry.getNet());
		}
		FileOutputStream out;
		out = new FileOutputStream(xlsPath);
		workbook.write(out);
		out.close();
		
	}

}
