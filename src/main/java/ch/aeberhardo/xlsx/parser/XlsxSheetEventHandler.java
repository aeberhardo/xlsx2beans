package ch.aeberhardo.xlsx.parser;

import java.util.Date;

public interface XlsxSheetEventHandler {

	void startRow(int rowNum);

	void endRow(int rowNum);

	void stringCell(int rowNum, int colIndex, String colName, String cellValue);

	void doubleCell(int rowNum, int colIndex, String colName, Double cellValue);

	void dateCell(int rowNum, int colIndex, String colName, Date cellValue);

}
