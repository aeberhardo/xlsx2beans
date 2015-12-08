package ch.aeberhardo.xlsx2beans;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ch.aeberhardo.xlsx2beans.converter.XlsxBeanConverterException;
import ch.aeberhardo.xlsx2beans.converter.XlsxToBeanConverterHandler;
import ch.aeberhardo.xlsx2beans.parser.XlsxSheetParser;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * This class converts a XLSX spread sheet to a list of Java beans.
 * Every row in the spread sheet gets converted to a single bean.
 * 
 * The first row in the spread sheet has to contain the names used for mapping columns to Java bean properties.
 * The column mapping from spread sheet to Java bean is configured by annotating the setter methods on the target bean class with the {@link ch.aeberhardo.xlsx2beans.converter.XlsxColumnName} annotation.
 */
public class XlsxToBeanConverter {

	/**
	 * Parse a XLSX spread sheet stream and convert rows to a list of beans.
	 * The first row of the sheet has to contain the names of the bean properties to be mapped.
	 * Mapping is done by applying annotation on the target bean class.
	 * 
	 * @param xlsxWorkbookInputStream The XLSX spread sheet input stream to be converted.
	 * @param sheetIndex Index of the spread sheet tab starting with 0.
	 * @param beanType The type of class the spread sheet rows get converted to.
	 * @return A list of newly created beans. Every spread sheet row gets converted to a single bean instance.
	 */
	public <T> List<T> convert(InputStream xlsxWorkbookInputStream, int sheetIndex, Class<T> beanType) {

		try (OPCPackage pkg = OPCPackage.open(xlsxWorkbookInputStream)) {

			Workbook wb = new XSSFWorkbook(pkg);
			Sheet sheet = wb.getSheetAt(sheetIndex);

			return convert(sheet, beanType);

		} catch (InvalidFormatException | IOException e) {
			throw new XlsxBeanConverterException(e);
		}
	}
        
        
        /* supporting xls conversion */ 
        public <T> List<T> convertXls(InputStream xlsWorkbookInputStream, int sheetIndex, Class<T> beanType) {

		try {
			Workbook wb = new HSSFWorkbook(xlsWorkbookInputStream);
			Sheet sheet = wb.getSheetAt(sheetIndex);

			return convert(sheet, beanType);

		} catch (IOException e) {
			throw new XlsxBeanConverterException(e);
		}
	}

	/**
	 * Parse a XLSX spread sheet stream and convert rows to a list of beans.
	 * The first row of the sheet has to contain the names of the bean properties to be mapped.
	 * Mapping is done by applying annotation on the target bean class.
	 * 
	 * @param sheet The spread sheet to be converted.
	 * @param beanType The type of class the spread sheet rows get converted to.
	 * @return A list of newly created beans. Every spread sheet row gets converted to a single bean instance.
	 */
	public <T> List<T> convert(Sheet sheet, Class<T> beanType) {

		XlsxToBeanConverterHandler<T> handler = new XlsxToBeanConverterHandler<T>(beanType);
		XlsxSheetParser parser = new XlsxSheetParser();

		parser.parse(sheet, handler);

		return handler.getBeans();
	}

}
