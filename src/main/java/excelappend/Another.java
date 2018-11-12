package excelappend;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Another {

	private HSSFWorkbook workbookOld = null;
	private XSSFWorkbook workbookNew = null;
	private int lastColumn = 0;
	private HashMap<Integer, XSSFCellStyle> styleMap = new HashMap<Integer, XSSFCellStyle>();

	public void transform(String inpFn, String outFn) throws FileNotFoundException, IOException {
		XSSFSheet sheetNew;
		HSSFSheet sheetOld;
		
		workbookOld = new HSSFWorkbook(new BufferedInputStream(new FileInputStream(inpFn)));
		workbookNew = new XSSFWorkbook();
		
		this.workbookNew.setForceFormulaRecalculation(this.workbookOld.getForceFormulaRecalculation());

		this.workbookNew.setMissingCellPolicy(this.workbookOld.getMissingCellPolicy());

		for (int i = 0; i < this.workbookOld.getNumberOfSheets(); i++) {
			sheetOld = this.workbookOld.getSheetAt(i);

			sheetNew = this.workbookNew.createSheet(sheetOld.getSheetName());
			this.transform(sheetOld, sheetNew);
		}
		
		BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(outFn));
		workbookNew.write(out);

	}
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		Another an = new Another();
		an.transform("D:\\old.xls", "D:\\new.xlsx");
	}

	private void transform(HSSFSheet sheetOld, XSSFSheet sheetNew) {

		sheetNew.setDisplayFormulas(sheetOld.isDisplayFormulas());
		sheetNew.setDisplayGridlines(sheetOld.isDisplayGridlines());
		sheetNew.setDisplayGuts(sheetOld.getDisplayGuts());
		sheetNew.setDisplayRowColHeadings(sheetOld.isDisplayRowColHeadings());
		sheetNew.setDisplayZeros(sheetOld.isDisplayZeros());
		sheetNew.setFitToPage(sheetOld.getFitToPage());
		sheetNew.setForceFormulaRecalculation(sheetOld.getForceFormulaRecalculation());
		sheetNew.setHorizontallyCenter(sheetOld.getHorizontallyCenter());
		sheetNew.setMargin(Sheet.BottomMargin, sheetOld.getMargin(Sheet.BottomMargin));
		sheetNew.setMargin(Sheet.FooterMargin, sheetOld.getMargin(Sheet.FooterMargin));
		sheetNew.setMargin(Sheet.HeaderMargin, sheetOld.getMargin(Sheet.HeaderMargin));
		sheetNew.setMargin(Sheet.LeftMargin, sheetOld.getMargin(Sheet.LeftMargin));
		sheetNew.setMargin(Sheet.RightMargin, sheetOld.getMargin(Sheet.RightMargin));
		sheetNew.setMargin(Sheet.TopMargin, sheetOld.getMargin(Sheet.TopMargin));
		sheetNew.setPrintGridlines(sheetNew.isPrintGridlines());
		sheetNew.setRightToLeft(sheetNew.isRightToLeft());
		sheetNew.setRowSumsBelow(sheetNew.getRowSumsBelow());
		sheetNew.setRowSumsRight(sheetNew.getRowSumsRight());
		sheetNew.setVerticallyCenter(sheetOld.getVerticallyCenter());

		XSSFRow rowNew;
		for (Row row : sheetOld) {
			rowNew = sheetNew.createRow(row.getRowNum());
			if (rowNew != null)
				this.transform((HSSFRow) row, rowNew);
		}

		for (int i = 0; i < this.lastColumn; i++) {
			sheetNew.setColumnWidth(i, sheetOld.getColumnWidth(i));
			sheetNew.setColumnHidden(i, sheetOld.isColumnHidden(i));
		}

		for (int i = 0; i < sheetOld.getNumMergedRegions(); i++) {
			CellRangeAddress merged = sheetOld.getMergedRegion(i);
			sheetNew.addMergedRegion(merged);
		}
		
	
	}

	private void transform(HSSFRow rowOld, XSSFRow rowNew) {
		XSSFCell cellNew;
		rowNew.setHeight(rowOld.getHeight());
		if (rowOld.getRowStyle() != null) {
			Integer hash = rowOld.getRowStyle().hashCode();
			if (!this.styleMap.containsKey(hash))
				this.transform(hash, rowOld.getRowStyle(), this.workbookNew.createCellStyle());
			rowNew.setRowStyle(this.styleMap.get(hash));
		}
		for (Cell cell : rowOld) {
			cellNew = rowNew.createCell(cell.getColumnIndex(), cell.getCellType());
			if (cellNew != null)
				this.transform((HSSFCell) cell, cellNew);
		}
		this.lastColumn = Math.max(this.lastColumn, rowOld.getLastCellNum());
	}

	private void transform(HSSFCell cellOld, XSSFCell cellNew) {
		cellNew.setCellComment(cellOld.getCellComment());

		Integer hash = cellOld.getCellStyle().hashCode();
		if (!this.styleMap.containsKey(hash)) {
			this.transform(hash, cellOld.getCellStyle(),
					this.workbookNew.createCellStyle());
		}
		cellNew.setCellStyle(this.styleMap.get(hash));

		switch (cellOld.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			cellNew.setCellValue(cellOld.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			cellNew.setCellValue(cellOld.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			cellNew.setCellValue(cellOld.getCellFormula());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			cellNew.setCellValue(cellOld.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING:
			cellNew.setCellValue(cellOld.getStringCellValue());
			break;
		default:
			System.out.println("transform: Unbekannter Zellentyp "
					+ cellOld.getCellType());
		}
	}

	private void transform(Integer hash, HSSFCellStyle styleOld, XSSFCellStyle styleNew) {
		styleNew.setAlignment(styleOld.getAlignment());
		styleNew.setBorderBottom(styleOld.getBorderBottom());
		styleNew.setBorderLeft(styleOld.getBorderLeft());
		styleNew.setBorderRight(styleOld.getBorderRight());
		styleNew.setBorderTop(styleOld.getBorderTop());
		styleNew.setDataFormat(this.transform(styleOld.getDataFormat()));
		styleNew.setFillBackgroundColor(styleOld.getFillBackgroundColor());
		styleNew.setFillForegroundColor(styleOld.getFillForegroundColor());
		styleNew.setFillPattern(styleOld.getFillPattern());
		styleNew.setFont(this.transform(styleOld.getFont(this.workbookOld)));
		styleNew.setHidden(styleOld.getHidden());
		styleNew.setIndention(styleOld.getIndention());
		styleNew.setLocked(styleOld.getLocked());
		styleNew.setVerticalAlignment(styleOld.getVerticalAlignment());
		styleNew.setWrapText(styleOld.getWrapText());
		this.styleMap.put(hash, styleNew);
	}

	private short transform(short index) {
		DataFormat formatOld = this.workbookOld.createDataFormat();
		DataFormat formatNew = this.workbookNew.createDataFormat();
		return formatNew.getFormat(formatOld.getFormat(index));
	}

	private XSSFFont transform(HSSFFont fontOld) {
		XSSFFont fontNew = this.workbookNew.createFont();
		fontNew.setBoldweight(fontOld.getBoldweight());
		fontNew.setCharSet(fontOld.getCharSet());
		fontNew.setColor(fontOld.getColor());
		fontNew.setFontName(fontOld.getFontName());
		fontNew.setFontHeight(fontOld.getFontHeight());
		fontNew.setItalic(fontOld.getItalic());
		fontNew.setStrikeout(fontOld.getStrikeout());
		fontNew.setTypeOffset(fontOld.getTypeOffset());
		fontNew.setUnderline(fontOld.getUnderline());
		return fontNew;
	}
}