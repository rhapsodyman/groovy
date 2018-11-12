package excelappend;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Transformer {
	
	public static void main(String[] args) throws InvalidFormatException, IOException {
		Transformer tr = new Transformer();
		tr.transform("D:\\old.xls", "D:\\new.xlsx");
	}
	

	public  void transform(String inpFn, String outFn) throws InvalidFormatException, 	IOException {

		InputStream in = new BufferedInputStream(new FileInputStream(inpFn));
		OutputStream out = null;
		Workbook wbIn = null, wbOut = null;
		try {
			wbIn = new HSSFWorkbook(in);
			File outF = new File(outFn);
			if (outF.exists())
				outF.delete();

			wbOut = new XSSFWorkbook();
			int sheetCnt = wbIn.getNumberOfSheets();
			for (int sheetInd = 0; sheetInd < sheetCnt; sheetInd++) {
				Sheet sIn = wbIn.getSheetAt(sheetInd);
				Sheet sOut = wbOut.createSheet(sIn.getSheetName());
				Iterator<Row> rowIt = sIn.rowIterator();
				while (rowIt.hasNext()) {
					Row rowIn = rowIt.next();
					Row rowOut = sOut.createRow(rowIn.getRowNum());

					Iterator<Cell> cellIt = rowIn.cellIterator();
					while (cellIt.hasNext()) {
						Cell cellIn = cellIt.next();
						Cell cellOut = rowOut.createCell(
								cellIn.getColumnIndex(), cellIn.getCellType());

						switch (cellIn.getCellType()) {
						case Cell.CELL_TYPE_BLANK:
							break;

						case Cell.CELL_TYPE_BOOLEAN:
							cellOut.setCellValue(cellIn.getBooleanCellValue());
							break;

						case Cell.CELL_TYPE_ERROR:
							cellOut.setCellValue(cellIn.getErrorCellValue());
							break;

						case Cell.CELL_TYPE_FORMULA:
							cellOut.setCellFormula(cellIn.getCellFormula());
							break;

						case Cell.CELL_TYPE_NUMERIC:
							cellOut.setCellValue(cellIn.getNumericCellValue());
							break;

						case Cell.CELL_TYPE_STRING:
							cellOut.setCellValue(cellIn.getStringCellValue());
							break;
						}

						CellStyle styleIn = cellIn.getCellStyle();
						CellStyle styleOut = cellOut.getCellStyle();
						styleOut.setDataFormat(styleIn.getDataFormat());
						cellOut.setCellComment(cellIn.getCellComment());
						styleOut.setFillBackgroundColor(styleIn.getFillBackgroundColor());

					
					}
				}
			}
			out = new BufferedOutputStream(new FileOutputStream(outF));
			wbOut.write(out);

		} finally {
			in.close();
			out.close();
			wbOut.close();
			wbIn.close();
		}
	}
}
