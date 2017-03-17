package accounting.shell.services;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import accounting.shell.models.Settlement;


public class SunocoXls {
	public static void processSettlementXls(String path, List<List<Settlement>> lists) throws IOException {
		FileInputStream input = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = null;
		List<Settlement> settlements = null;
		int coeff;
		int bblCol;
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			// find corresponding list for each tab
			String sheetName = workbook.getSheetName(i).toLowerCase().trim();	
			sheet = workbook.getSheetAt(i);
			if (sheetName.contains("bulk") && sheetName.contains("ar")) {
				settlements = lists.get(0);
				coeff = -1;
				bblCol = 3;
			} else if (sheetName.contains("lease") && sheetName.contains("ar")) {
				settlements = lists.get(1);
				coeff = -1;
				bblCol = 3;
			} else if (sheetName.contains("bulk") && sheetName.contains("ap")) {
				settlements = lists.get(2);
				coeff = 1;
				bblCol = 2;
			} else {
				continue;
			}	
			
			double volEps = 0.05;
			int rowId;
			Map<Integer, Boolean> volMap = getVolMap(sheet, bblCol);
			boolean[] mapped = new boolean[settlements.size()];
			for (rowId = 6; rowId <= sheet.getLastRowNum(); rowId++)
	    	{
				Row row = sheet.getRow(rowId);
				if (row.getCell(1) == null) {
					break;
				} else if (row.getCell(1).getStringCellValue() == "") {
					break;
				}
				if (row.getCell(bblCol) == null || row.getCell(bblCol + 1) == null) {
					continue;
				}
				
				double volume = row.getCell(bblCol).getNumericCellValue();
				double price = row.getCell(bblCol + 1).getNumericCellValue();
				
				
				double minDelta = Double.MAX_VALUE;
				int minIndex = -1;
				for (int j = 0; j < settlements.size(); j++) {
					if (mapped[j] == true) continue;
					
					Settlement s = settlements.get(j);
					double shellVolume = s.getVolume();
					double shellPrice = s.getSettleAmount();
					
					// only one volume entry
					if (volMap.get(rowId)) {
						if (Math.abs(shellVolume - volume) < volEps) {
							row.getCell(bblCol + 4).setCellValue(shellVolume);
							row.getCell(bblCol + 5).setCellValue(shellPrice * coeff);
							double base = shellPrice * coeff / shellVolume;
							row.getCell(bblCol + 6).setCellValue(base);
							mapped[j] = true;
							break;
						}
					}
					// multiple volume entries
					else {
						if (Math.abs(shellVolume - volume) > volEps) {
							continue;
						}
						double delta = Math.abs(coeff * shellPrice - price);
						if (delta < minDelta) {
							minDelta = delta;
							minIndex = j;
						}
					}
				}
				
				// found entry with closest price for same volume
				if (minIndex != -1) {
					Settlement sett = settlements.get(minIndex);
					row.getCell(bblCol + 4).setCellValue(sett.getVolume());
					row.getCell(bblCol + 5).setCellValue(sett.getSettleAmount() * coeff);
					double base = sett.getSettleAmount() * coeff / sett.getVolume();
					row.getCell(bblCol + 6).setCellValue(base);
					mapped[minIndex] = true;
				}
	    	}
			
			// skip the summary line
			rowId += 30;
			// If not all matched, append leftover at the bottom
			for (int j = 0; j < settlements.size(); j++) {
				if (mapped[j] == true) continue;
				Settlement sett = settlements.get(j);
				Row row = sheet.createRow(rowId++);
				Cell cell = row.createCell(bblCol + 4);
				cell.setCellValue(sett.getVolume());
				cell = row.createCell(bblCol + 5);
				cell.setCellValue(sett.getSettleAmount() * coeff);
				cell = row.createCell(bblCol + 6);
				cell.setCellValue(sett.getSettleAmount() * coeff / sett.getVolume());
			}
		}
		input.close();
		
		//update xls
		FileOutputStream outFile =new FileOutputStream(path);
        workbook.write(outFile);
        workbook.close();
        outFile.close();
	}

	
	private static Map<Integer, Boolean> getVolMap(XSSFSheet sheet, int volCol) {
		Map<String, Integer> countMap = new HashMap<String, Integer>();
		int lastRow = sheet.getLastRowNum();
		for (int i = 6; i < lastRow; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				lastRow = i;
				break;
			}
			if (row.getCell(volCol) == null) {
				lastRow = i;
				break;
			}
			double vol = row.getCell(volCol).getNumericCellValue();
			String volStr = String.format ("%.2f", vol);
			if (countMap.containsKey(volStr)) {
				countMap.put(volStr, countMap.get(volStr) + 1);
			} else {
				countMap.put(volStr, 1);
			}
		}
		
		// find duplicates true for singlet, false for duplicates
		Map<Integer, Boolean> map = new HashMap<Integer, Boolean>();
		for (int i = 6; i < lastRow; i++) {
			Row row = sheet.getRow(i);
			double vol = row.getCell(volCol).getNumericCellValue();
			String volStr = String.format ("%.2f", vol);
			if (countMap.get(volStr) == 1) map.put(i, true);
			else map.put(i, false);
		}
		return map;
	}
}
