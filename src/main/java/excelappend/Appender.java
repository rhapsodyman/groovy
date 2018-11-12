package excelappend;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Appender {
	  private static final Map<String, String> firstMap;
	  private static final Map<String, String> secondtMap;
	    static {
	    	firstMap = new HashMap<String, String>();
	    	firstMap.put("money", "one");
	    	firstMap.put("city", "two");
	    	firstMap.put("country", "tree");
	    	firstMap.put("state", "four");
	    	
	    	secondtMap = new HashMap<String, String>();
	    	secondtMap.put("bfd", "five");
	    	secondtMap.put("city", "six");
	    	secondtMap.put("vad", "seven");
	    	secondtMap.put("state", "eight");
	    	
	    }
	    
	    public static void main(String[] args) throws Exception {
	    	Appender ap = new Appender();
	    	ap.doIt();
	    }

	public void doIt() throws IOException{

		Sheet sheet = null;
		Workbook workbook = null;

		String filePath = "D:\\document.xlsx";
		FileInputStream inputStream = new FileInputStream(new File(filePath));

		if (filePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook(inputStream);
		}

		else if (filePath.endsWith("xls")) {
			workbook = new HSSFWorkbook(inputStream);
		}
		
		inputStream.close();

		sheet = workbook.getSheetAt(0);
		
		int lastRowNum = sheet.getLastRowNum() + 1;
		Row createRow = sheet.createRow(lastRowNum);
		Row firstRow = sheet.getRow(0);
		
		for (Entry<String, String> mapEntry : firstMap.entrySet()) {
			// try to find a cell in the first row
			
			int index = findCellColumnIndex(firstRow, mapEntry.getKey());
			
			if ( index != -1){
				Cell createCell = createRow.createCell(index);
				createCell.setCellValue(mapEntry.getValue());
			}
			else{
				index = firstRow.getLastCellNum();
				Cell firstowCell = firstRow.createCell(index);
				firstowCell.setCellValue(mapEntry.getKey());
				
				Cell currentCell = createRow.createCell(index);
				currentCell.setCellValue(mapEntry.getValue());
			}
		}
		
		
		
		
		
		//==============
		lastRowNum = sheet.getLastRowNum() + 1;
		 createRow = sheet.createRow(lastRowNum);
		firstRow = sheet.getRow(0);
		
		for (Entry<String, String> mapEntry : secondtMap.entrySet()) {
			// try to find a cell in the first row
			
			int index = findCellColumnIndex(firstRow, mapEntry.getKey());
			
			if ( index != -1){
				Cell createCell = createRow.createCell(index);
				createCell.setCellValue(mapEntry.getValue());
			}
			else{
				index = firstRow.getLastCellNum();
				Cell firstCell = firstRow.createCell(index);
				firstCell.setCellValue(mapEntry.getKey());
				
				Cell currentCell = createRow.createCell(index);
				currentCell.setCellValue(mapEntry.getValue());
			}
		}
		
		
		
		
		

		 //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(filePath));
        workbook.write(out);
        
        workbook.close();
       out.close();

	}
	
	public int findCellColumnIndex(Row row, String cellName){
		for (Cell cell : row) {
			if (cell != null && cell.getStringCellValue() != null && cell.getStringCellValue().equals(cellName)) {
				return cell.getColumnIndex();
			}
		}
		return -1;
		
	}
}
