package ch.aeberhardo.xlsx2beans.parser;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.Format;
import java.text.ParseException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class XlsxSheetParser {

    private FormulaEvaluator m_formulaEvaluator;
    private DataFormatter m_dataFormatter;

    public void parse(Sheet sheet, XlsxSheetEventHandler handler) {
        prepareParser(sheet);
        parseRows(sheet, handler);
    }

    private void prepareParser(Sheet sheet) {
        m_formulaEvaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        m_dataFormatter = new DataFormatter();
    }

    private void parseRows(Sheet sheet, XlsxSheetEventHandler handler) {

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
                throw new XlsxParserException("Error while parsing header (rowNum=" + row.getRowNum() + ", colIndex=" + cell.getColumnIndex() + "): "
                        + e.getMessage(), e);
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
                throw new XlsxParserException("Error while parsing cell (rowNum=" + row.getRowNum() + ", colIndex=" + cell.getColumnIndex()
                        + "): No header name defined!");
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

    /**
     * Handle a numeric cell. Numeric cells can contain numbers or dates.
     *
     * @param cell
     * @param rowNum
     * @param colIndex
     * @param colName
     * @param handler
     */
    private void handleNumericCell(Cell cell, int rowNum, int colIndex, String colName, XlsxSheetEventHandler handler) {

        if (DateUtil.isCellDateFormatted(cell)) {
            handleDateCell(cell, rowNum, colIndex, colName, handler);

        } else {
            handleNumberCell(cell, rowNum, colIndex, colName, handler);
        }

    }

    private void handleDateCell(Cell cell, int rowNum, int colIndex, String colName, XlsxSheetEventHandler handler) {
        handler.dateCell(rowNum, colIndex, colName, cell.getDateCellValue());
    }

    /**
     * Parse a cell containing a number. The number in the spreadsheet cell will
     * be converted to a BigDecimal to avoid precision loss. It is possible for
     * the the number in the spreadsheet to have a custom format, e.g.
     * "#,##0.00000".
     *
     * @param cell
     * @param rowNum
     * @param colIndex
     * @param colName
     * @param handler
     */
    private void handleNumberCell(Cell cell, int rowNum, int colIndex, String colName, XlsxSheetEventHandler handler) {

        String pattern = getDecimalFormatPattern(cell);

        String formattedCellValueAsString = getFormattedCellValueAsString(cell);

        try {
            if (pattern == null) {
                //fallback where as no pattern has been defined on cell. FIXME?
                if (isNumeric(formattedCellValueAsString)) {
                    handler.numberCell(rowNum, colIndex, colName, BigDecimal.valueOf(Double.valueOf(formattedCellValueAsString)));
                } else {
                    handler.stringCell(rowNum, colIndex, colName, formattedCellValueAsString);
                }
            } else {
                BigDecimal value = parseBigDecimal(formattedCellValueAsString, pattern);
                handler.numberCell(rowNum, colIndex, colName, value);
            }
        } catch (ParseException e) {
            throw new XlsxParserException("Error while parsing cell (rowNum=" + rowNum + ", colIndex=" + colIndex + ")!", e);
        }
    }

    private String getDecimalFormatPattern(Cell cell) {

        Format format = m_dataFormatter.createFormat(cell);

        if (!(format instanceof DecimalFormat)) {
            if (format == null) {
                return null;
            }
            throw new XlsxParserException("Error while parsing cell (rowNum=" + cell.getRow().getRowNum() + ", colIndex=" + cell.getColumnIndex()
                    + "): Format not supported! Expected " + DecimalFormat.class.getSimpleName() + " but got " + format.getClass().getSimpleName() + ".");
        }

        return ((DecimalFormat) format).toPattern();
    }

    private String getFormattedCellValueAsString(Cell cell) {
        return m_dataFormatter.formatCellValue(cell, m_formulaEvaluator);
    }

    /**
     * Parse a string and convert to a BigDecimal according to a formatting
     * pattern. Parsing uses BigDecimals to prevent precision loss.
     *
     * @param numberAsString
     * @param pattern
     * @return
     * @throws ParseException
     */
    private BigDecimal parseBigDecimal(String numberAsString, String pattern) throws ParseException {
        DecimalFormat decimalFormat = new DecimalFormat(pattern);
        decimalFormat.setParseBigDecimal(true);
        return (BigDecimal) decimalFormat.parse(numberAsString);
    }

    /**
     * Returns the actual type of the cell. If the cell contains a formula, the
     * resulting cell type is returned.
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

    public static boolean isNumeric(String str) {
        try {
            double d = Double.parseDouble(str);
        } catch (NumberFormatException nfe) {
            return false;
        }
        return true;
    }

}
