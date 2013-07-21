package ch.aeberhardo.xlsx2beans.parser;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class XlsxSheetParser {

	public void parse(XSSFSheet sheet, XlsxSheetEventHandler handler) {
		parseRows(sheet, handler);
	}

	private void parseRows(XSSFSheet sheet, XlsxSheetEventHandler handler) {

		Map<Integer, String> headerMap = parseHeaderCells(sheet.getRow(0));

		for (Row row : sheet) {

			if (row.getRowNum() == 0) {
				continue;
			}

			handler.startRow(row.getRowNum());

			parseCells(row, headerMap, handler);

			handler.endRow(row.getRowNum());
		}

	}

	private Map<Integer, String> parseHeaderCells(Row row) {

		Map<Integer, String> headerMap = new HashMap<>();

		for (Cell cell : row) {

			try {

				int colNum = cell.getColumnIndex();
				String headerCellName = cell.getStringCellValue();

				validateHeaderCellNameUnique(headerMap, headerCellName);

				headerMap.put(colNum, headerCellName);

			} catch (Exception e) {
				throw new XlsxParserException("Error while parsing header (rowNum=" + row.getRowNum() + ", colIndex=" + cell.getColumnIndex() + "): " + e.getMessage(), e);
			}
		}

		return headerMap;

	}

	private void validateHeaderCellNameUnique(Map<Integer, String> headerMap, String name) {

		if (headerMap.containsValue(name)) {
			throw new IllegalStateException("The column header with name '" + name + "' is not unique!");
		}

	}

	private void parseCells(Row row, Map<Integer, String> headerMap, XlsxSheetEventHandler handler) {

		for (Cell cell : row) {

			int colIndex = cell.getColumnIndex();
			String colName = headerMap.get(colIndex);

			if (colName == null || colName.isEmpty()) {
				throw new XlsxParserException("Error while parsing cell (rowNum=" + row.getRowNum() + ", colIndex=" + cell.getColumnIndex() + "): No header name defined!");
			}

			handleCell(cell, colName, handler);
		}

	}

	private void handleCell(Cell cell, String colName, XlsxSheetEventHandler handler) {

		int cellType = getActualCellType(cell);
		int rowNum = cell.getRowIndex();
		int colIndex = cell.getColumnIndex();

		if (cellType == Cell.CELL_TYPE_STRING) {
			handleStringCell(cell, rowNum, colIndex, colName, handler);

		} else if (cellType == Cell.CELL_TYPE_NUMERIC) {
			handleNumericCell(cell, rowNum, colIndex, colName, handler);
		}

	}

	private void handleStringCell(Cell cell, int rowNum, int colIndex, String colName, XlsxSheetEventHandler handler) {
		handler.stringCell(rowNum, colIndex, colName, cell.getStringCellValue());
	}

	private void handleNumericCell(Cell cell, int rowNum, int colIndex, String colName, XlsxSheetEventHandler handler) {

		if (DateUtil.isCellDateFormatted(cell)) {
			handler.dateCell(rowNum, colIndex, colName, cell.getDateCellValue());

		} else {
			handler.doubleCell(rowNum, colIndex, colName, cell.getNumericCellValue());
		}

	}
	
	/**
	 * Returns the actual type of the cell.
	 * If the cell contains a formula, the resulting cell type is returned.
	 * 
	 * @param cell
	 * @return Type of the cell after evaluating formulas.
	 */
	private int getActualCellType(Cell cell) {

		int cellType = cell.getCellType();

		if (cellType == Cell.CELL_TYPE_FORMULA) {
			return cell.getCachedFormulaResultType();

		} else {
			return cellType;
		}

	}	

}
