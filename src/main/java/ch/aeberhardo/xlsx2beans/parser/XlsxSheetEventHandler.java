package ch.aeberhardo.xlsx2beans.parser;

import java.util.Date;

public interface XlsxSheetEventHandler {

	void startRow(int rowNum);

	void endRow(int rowNum);

	void stringCell(int rowNum, int colIndex, String colName, String cellValue);

	void numberCell(int rowNum, int colIndex, String colName, Number cellValue);

	void dateCell(int rowNum, int colIndex, String colName, Date cellValue);

}
