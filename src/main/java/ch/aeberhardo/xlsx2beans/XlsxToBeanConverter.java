package ch.aeberhardo.xlsx2beans;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ch.aeberhardo.xlsx2beans.converter.XlsxBeanConverterException;
import ch.aeberhardo.xlsx2beans.converter.XlsxToBeanConverterHandler;
import ch.aeberhardo.xlsx2beans.parser.XlsxSheetParser;

public class XlsxToBeanConverter {

	public <T> List<T> convert(InputStream xlsxWorkbookInputStream, int sheetIndex, Class<T> beanType) {

		try (OPCPackage pkg = OPCPackage.open(xlsxWorkbookInputStream)) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(sheetIndex);

			return convert(sheet, beanType);

		} catch (InvalidFormatException | IOException e) {
			throw new XlsxBeanConverterException(e);
		}
	}

	public <T> List<T> convert(XSSFSheet sheet, Class<T> beanType) {

		XlsxToBeanConverterHandler<T> handler = new XlsxToBeanConverterHandler<T>(beanType);
		XlsxSheetParser parser = new XlsxSheetParser();

		parser.parse(sheet, handler);

		return handler.getBeans();
	}

}
