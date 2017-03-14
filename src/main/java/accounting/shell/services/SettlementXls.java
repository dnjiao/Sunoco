package accounting.shell.services;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import accounting.shell.models.Settlement;


public class SettlementXls {
	//final static Logger logger = Logger.getLogger(SettlementXls.class);
	
	public static List<List<Settlement>> readShellXls(String xlsPath) throws Exception {
		File inFile = new File(xlsPath);
		Workbook workbook = new XSSFWorkbook(new FileInputStream(inFile));
		if (findDataSheet(workbook) == -1) {
			throw new RuntimeException("Data sheet is not found in " + xlsPath);
		}
		
		Sheet sheet = workbook.getSheetAt(findDataSheet(workbook));
		Row row = sheet.getRow(0);
		Cell cell = null;
		Map<String, Integer> colNameMap = getColNameMap(row);
		List<Settlement> bulkSA = new ArrayList<Settlement>();
		List<Settlement> bulkPA = new ArrayList<Settlement>();
		List<Settlement> leaseSA = new ArrayList<Settlement>();
		List<Settlement> leasePA = new ArrayList<Settlement>();
		for (int rowId = 1; rowId <= sheet.getLastRowNum(); rowId++)
    	{
			row = sheet.getRow(rowId);
			if (row == null) continue;
			Settlement settlement = new Settlement();
			cell = row.getCell(colNameMap.get("period"));
			if (cell != null) {
				if (cell.getStringCellValue() == "") continue;
				settlement.setPeriod(cell.getStringCellValue());
			} else 
				break;
			cell = row.getCell(colNameMap.get("flag"));
			if (cell != null) {
				settlement.setBuySellFlag(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("status"));
			if (cell != null) {
				settlement.setStatus(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("date"));
			if (cell != null) {
				DateFormat df = new SimpleDateFormat("M/dd/yyyy");
				settlement.setDate(df.format(cell.getDateCellValue()));
			}
			cell = row.getCell(colNameMap.get("contract"));
			if (cell != null) {
				settlement.setContractNo((int)cell.getNumericCellValue());
			}
			cell = row.getCell(colNameMap.get("smart"));
			if (cell != null) {
				settlement.setSmartNo(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("commit"));
			if (cell != null) {
				settlement.setCommitment((int)cell.getNumericCellValue());
			}
			cell = row.getCell(colNameMap.get("dealtrack"));
			if (cell != null) {
				settlement.setDealTrackNo((int)cell.getNumericCellValue());
			}
			cell = row.getCell(colNameMap.get("volume"));
			if (cell != null) {
				settlement.setVolume(cell.getNumericCellValue());
			}
			cell = row.getCell(colNameMap.get("unit"));
			if (cell != null) {
				settlement.setUnit(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("price"));
			if (cell != null) {
				settlement.setPrice(cell.getNumericCellValue());
			}
			cell = row.getCell(colNameMap.get("currency"));
			if (cell != null) {
				settlement.setCurrency(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("type"));
			if (cell != null) {
				settlement.setCashFlowType(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("settleamount"));
			if (cell != null) {
				settlement.setSettleAmount(cell.getNumericCellValue());
			}
			cell = row.getCell(colNameMap.get("location"));
			if (cell != null) {
				settlement.setLocation(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("leaseno"));
			if (cell != null) {
				settlement.setLeaseNo(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("leasename"));
			if (cell != null) {
				settlement.setLeaseName(cell.getStringCellValue());
			}			
			cell = row.getCell(colNameMap.get("product"));
			if (cell != null) {
				settlement.setProduct(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("pipeline"));
			if (cell != null) {
				settlement.setPipeline(cell.getStringCellValue());
			}
			cell = row.getCell(colNameMap.get("agreement"));
			if (cell != null) {
				settlement.setAgreement(cell.getStringCellValue());
			}
			
			// Determine category based on buy/sell flag and lease name
			
			if (settlement.getBuySellFlag().equalsIgnoreCase("Buy")) {
				if (settlement.getLeaseName() == null || 
						settlement.getLeaseName().equalsIgnoreCase("NA") ||
						settlement.getLeaseName().equalsIgnoreCase("N/A") ||
						settlement.getLeaseName() == "") {
					bulkSA.add(settlement);
				} else {
					leaseSA.add(settlement);
				}
			} else if (settlement.getBuySellFlag().equalsIgnoreCase("Sell")) {
				if (settlement.getLocation().endsWith("Lease Sale")) {
					leasePA.add(settlement);
				} else {
					bulkPA.add(settlement);
				}
			} else {
				System.out.println("ERROR in row " + Integer.toString(rowId) + " invalid value in Buy/Sell Flag column.");
			}
    	}
		List<Settlement> bsList = groupingSum(bulkSA);
		List<Settlement> lsList = groupingSum(leaseSA);
		List<Settlement> bpList = groupingSum(bulkPA);
		List<Settlement> lpList = groupingSum(leasePA);
		
		List<List<Settlement>> masterList = new ArrayList<List<Settlement>>(4);
		masterList.add(bsList);
		masterList.add(lsList);
		masterList.add(bpList);
		masterList.add(lpList);
		workbook.close();
		return masterList;
	}
	
	private static int findDataSheet(Workbook workbook) {
		int index = -1;
		for (int i = 0; i < workbook.getNumberOfSheets(); i++)
		{
		    Sheet sheet = workbook.getSheetAt(i);
		    
		    Row row = sheet.getRow(0);
		    if (row == null) continue;
		    
		    Cell cell = row.getCell(0);
		    if (cell == null) continue;
		    
			String name = cell.getStringCellValue().toLowerCase();
			if (name.contains("production") && name.contains("period")) {
				index = i;
				break;
			}
		}
		
		return index;
	}

	// write tabs to Shell Xls
	public static void writeShellXls(List<List<Settlement>> lists, String path) throws FileNotFoundException, IOException {
		
		Workbook workbook = new XSSFWorkbook(new FileInputStream(path));
		// set double cell format
		CellStyle doubleCellStyle = workbook.createCellStyle();
		doubleCellStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
		
		// set bold font format
		CellStyle boldStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		boldStyle.setFont(font);

		// Add 4 tabs in the original Shell settlement xls iteratively
		String[] tabs = {"Bulk AR", "Lease AR", "Bulk AP", "Royalty"};
		for (int i = 0; i < 4; i++) {
			if (workbook.getSheet(tabs[i]) != null) {
				continue;
			}
			
			List<Settlement> list = lists.get(i);	
			Sheet sheet = workbook.createSheet(tabs[i]);
			// write title row
			Row row = sheet.createRow(0);
			Cell cell = null;
			String[] cols = {"ProductionPeriod", "STA Netting Buy/Sell Flag", "Volume Status",
					"Event Date", "Contract#", "Smart#", "Commitment#", "DealTracking #", 
					"BAVVolume", "BAVVolume Unit", "DeliveryCurrency", "Cash FlowType", 
					"CurrentSettle Amount", "Location",	"Lease#", "Lease Name", 
					"Product", "Pipeline Name", "Master Netting Agreement"};
			for (int j = 0; j < cols.length; j++) {
				cell = row.createCell(j);
				cell.setCellValue(cols[j]);
				cell.setCellStyle(boldStyle);
			}
			// write records
			int rowId = 1;
			for (int k = 0; k < list.size(); k++) {
				Settlement sett = list.get(k);
				row = sheet.createRow(rowId++);
				cell = row.createCell(0);
				cell.setCellValue(sett.getPeriod());
				cell = row.createCell(1);
				cell.setCellValue(sett.getBuySellFlag());
				cell = row.createCell(2);
				cell.setCellValue(sett.getStatus());
				cell = row.createCell(3);
				cell.setCellValue(sett.getDate());
				cell = row.createCell(4);
		        cell.setCellValue(sett.getContractNo());
		        cell = row.createCell(5);
				cell.setCellValue(sett.getSmartNo());
				cell = row.createCell(6);
				cell.setCellValue(sett.getCommitment());
				cell = row.createCell(7);
				cell.setCellValue(sett.getDealTrackNo());
				cell = row.createCell(8);
				cell.setCellValue(sett.getVolume());
				cell.setCellStyle(doubleCellStyle);
				cell = row.createCell(9);
				cell.setCellValue(sett.getUnit());
				cell = row.createCell(10);
				cell.setCellValue(sett.getCurrency());
				cell = row.createCell(11);
				cell.setCellValue(sett.getCashFlowType());
				cell = row.createCell(12);
				cell.setCellValue(sett.getSettleAmount());
				cell.setCellStyle(doubleCellStyle);
				cell = row.createCell(13);
				cell.setCellValue(sett.getLocation());
				cell = row.createCell(14);
				cell.setCellValue(sett.getLeaseNo());
				cell = row.createCell(15);
				cell.setCellValue(sett.getLeaseName());
				cell = row.createCell(16);
				cell.setCellValue(sett.getProduct());
				cell = row.createCell(17);
				cell.setCellValue(sett.getPipeline());
				cell = row.createCell(18);
				cell.setCellValue(sett.getAgreement());
			}
		}
		
		FileOutputStream fileOut = new FileOutputStream(path);
	    workbook.write(fileOut);
	    workbook.close();
	    fileOut.close();
		
	}
	private static Map<String, Integer> getColNameMap(Row row) {
		Map<String, Integer> map = new HashMap<String, Integer>();
		short colTotal = row.getLastCellNum();
		for (int i = 0; i < colTotal; i++) {
			Cell cell =row.getCell(i);
			if (cell != null) {
				String name = cell.getStringCellValue().toLowerCase();
				if (name.contains("production") && name.contains("period"))
					map.put("period", i);
				if (name.contains("buy") && name.contains("sell")) 
					map.put("flag", i);
				if (name.contains("volume") && name.contains("status"))
					map.put("status", i);
				if (name.contains("event") && name.contains("date")) 
					map.put("date", i);
				if (name.contains("contract") && name.contains("#"))
					map.put("contract", i);
				if (name.contains("smart") && name.contains("#"))
					map.put("smart", i);
				if (name.contains("commitment")) 
					map.put("commit", i);
				if (name.contains("deal") && name.contains("tracking"))
					map.put("dealtrack", i);
				if (name.trim().endsWith("volume"))
					map.put("volume", i);
				if (name.contains("volume") && name.contains("unit"))
					map.put("unit", i);
				if (name.trim().endsWith("price"))
					map.put("price", i);
				if (name.contains("currency"))
					map.put("currency", i);
				if (name.contains("cash") && name.contains("type"))
					map.put("type", i);
				if (name.contains("current") && name.contains("settle") && name.contains("amount"))
					map.put("settleamount", i);
				if (name.contains("location"))
					map.put("location", i);
				if (name.contains("lease") && name.contains("#"))
					map.put("leaseno", i);
				if (name.contains("lease") && name.contains("name"))
					map.put("leasename", i);
				if (name.trim().endsWith("product"))
					map.put("product", i);
				if (name.contains("pipeline") && name.contains("name"))
					map.put("pipeline", i);		
				if (name.contains("netting") && name.contains("agreement")) {
					map.put("agreement", i);
				}
			}
		}
		return map;
	}
	
	private static List<Settlement> groupingSum(List<Settlement> list) {
		Map<Integer, List<Settlement>> map = new HashMap<Integer, List<Settlement>>();
		map = list.stream().collect(Collectors.groupingBy(Settlement::getDealTrackNo));
		List<Settlement> newList = new ArrayList<Settlement>();
		for (Map.Entry<Integer, List<Settlement>> entry : map.entrySet()) {
		    double total = 0.0;
		    Settlement newSett = null;
		    int copyFlag = 0;
		    for (Settlement sett : entry.getValue()) {
		    	if (copyFlag == 0) {
		    		newSett = copySettlement(sett);
		    		copyFlag = 1;
		    	}
		    	total += sett.getSettleAmount();
		    	newSett.setCashFlowType("Cmdity");
		    	newSett.setSettleAmount(total);
		    }
		    newList.add(newSett);
		}
		Collections.sort(newList, new Comparator<Settlement>() {
		    @Override
		    public int compare(Settlement s1, Settlement s2) {
		    	double epsilon = 0.001;
		    	if (Math.abs(s1.getVolume() - s2.getVolume()) < epsilon) {
		    		return Double.compare(s1.getSettleAmount(), s2.getSettleAmount());
		    	} 
		    	return Double.compare(s1.getVolume(), s2.getVolume());
		    }
		});
		return newList;
	}
	
	// partially copy Settlement object to new one
	private static Settlement copySettlement(Settlement sett) {
		Settlement newSett = new Settlement();
		newSett.setPeriod(sett.getPeriod());
		newSett.setBuySellFlag(sett.getBuySellFlag());
		newSett.setStatus(sett.getStatus());
		newSett.setDate(sett.getDate());
		newSett.setContractNo(sett.getContractNo());
		newSett.setSmartNo(sett.getSmartNo());
		newSett.setCommitment(sett.getCommitment());
		newSett.setDealTrackNo(sett.getDealTrackNo());
		newSett.setVolume(sett.getVolume());
		newSett.setUnit(sett.getUnit());
		newSett.setCurrency(sett.getCurrency());
		newSett.setLocation(sett.getLocation());
		newSett.setLeaseNo(sett.getLeaseNo());
		newSett.setLeaseName(sett.getLeaseName());
		newSett.setProduct(sett.getProduct());
		newSett.setPipeline(sett.getPipeline());
		return newSett;
	}


	
}

	
