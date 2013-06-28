package ch.aeberhardo.xlsx.parser;

import static org.mockito.Mockito.*;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import ch.aeberhardo.xlsx.parser.XlsxSheetEventHandler;
import ch.aeberhardo.xlsx.parser.XlsxSheetParser;

public class XlsxSheetParserTest {

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void test_handlerCalls() {

		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test1.xlsx"))) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);

			XlsxSheetParser parser = new XlsxSheetParser();

			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);

			parser.parse(sheet, handlerMock);

			verify(handlerMock).startRow(1);
			verify(handlerMock).startRow(2);
			
			verify(handlerMock).stringCell(1, 0, "MyString1", "ABC");
			verify(handlerMock).stringCell(1, 1, "MyString2", "This is my string 1");
			verify(handlerMock).doubleCell(1, 2, "MyInteger", 123.0d);
			verify(handlerMock).doubleCell(1, 3, "MyDouble", 7.89d);
			verify(handlerMock).dateCell(1, 4, "MyDate", new SimpleDateFormat("dd.MM.yyyy hh:mm:ss").parse("12.01.2013 14:16:23"));
			
			verify(handlerMock).stringCell(2, 0, "MyString1", "DEF");
			verify(handlerMock).stringCell(2, 1, "MyString2", "This is my string 2");
			verify(handlerMock).doubleCell(2, 2, "MyInteger", 456.0d);
			verify(handlerMock).doubleCell(2, 3, "MyDouble", 9.87d);
			verify(handlerMock).dateCell(2, 4, "MyDate", new SimpleDateFormat("dd.MM.yyyy hh:mm:ss").parse("20.06.2011 23:50:33"));

		} catch (ParseException | InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}

	}

}
