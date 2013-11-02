package ch.aeberhardo.xlsx2beans.parser;

import static org.junit.Assert.*;
import static org.mockito.Matchers.*;
import static org.mockito.Mockito.*;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

public class XlsxSheetParserTest {

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void test_handlerCalls() {

		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test-valid.xlsx"))) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);

			XlsxSheetParser parser = new XlsxSheetParser();

			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);

			parser.parse(sheet, handlerMock);

			verify(handlerMock).startRow(1);
			verify(handlerMock).startRow(2);

			verify(handlerMock).stringCell(1, 0, "MyString1", "ABC");
			verify(handlerMock).stringCell(1, 1, "MyString2", "This is my string 1");
			verify(handlerMock).numberCell(1, 2, "MyInteger", new BigDecimal("123"));
			verify(handlerMock).numberCell(1, 3, "MyDouble", new BigDecimal("7.89"));
			verify(handlerMock).dateCell(1, 4, "MyDate", new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").parse("12.01.2013 14:16:23"));

			verify(handlerMock).stringCell(2, 0, "MyString1", "DEF");
			verify(handlerMock).stringCell(2, 1, "MyString2", "This is my string 2");
			verify(handlerMock).numberCell(2, 2, "MyInteger", new BigDecimal("456"));
			verify(handlerMock).numberCell(2, 3, "MyDouble", new BigDecimal("9.87"));
			verify(handlerMock).dateCell(2, 4, "MyDate", new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").parse("20.06.2011 23:50:33"));

		} catch (ParseException | InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}

	}

	@Test
	public void test_emptyRow() {

		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test-valid-with_empty_row.xlsx"))) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);

			XlsxSheetParser parser = new XlsxSheetParser();

			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);

			parser.parse(sheet, handlerMock);

			// Test sheet contains 3 rows.
			verify(handlerMock, times(3)).startRow(anyInt());

			// Row 3 is empty.
			verify(handlerMock).startRow(1);
			verify(handlerMock).startRow(2);
			verify(handlerMock).startRow(4);

		} catch (InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}

	}

	@Test
	public void test_invalidHeader() {

		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test-invalid_header.xlsx"))) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);

			XlsxSheetParser parser = new XlsxSheetParser();

			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);

			parser.parse(sheet, handlerMock);
			fail("Expected exception was not thrown!");

		} catch (XlsxParserException e) {
			assertTrue(e.getMessage().startsWith("Error while parsing header (rowNum=0, colIndex=2):"));

		} catch (InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}

	}

	@Test
	public void test_nonUniqueHeader() {

		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test-invalid-non_unique_header.xlsx"))) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);

			XlsxSheetParser parser = new XlsxSheetParser();

			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);

			parser.parse(sheet, handlerMock);
			fail("Expected exception was not thrown!");

		} catch (XlsxParserException e) {
			assertEquals("Error while parsing header (rowNum=0, colIndex=3): The column header with name 'MyString2' is not unique!", e.getMessage());

		} catch (InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}

	}

	@Test
	public void test_undefinedHeader() {

		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test-invalid-header_not_definied.xlsx"))) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);

			XlsxSheetParser parser = new XlsxSheetParser();

			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);

			parser.parse(sheet, handlerMock);
			fail("Expected exception was not thrown!");

		} catch (XlsxParserException e) {
			assertEquals("Error while parsing cell (rowNum=1, colIndex=5): No header name defined!", e.getMessage());

		} catch (InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}

	}

	@Test
	public void test_doublePrecision() {

		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test-valid-double_precision.xlsx"))) {

			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);

			XlsxSheetParser parser = new XlsxSheetParser();

			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);

			parser.parse(sheet, handlerMock);

			verify(handlerMock).startRow(1);

			verify(handlerMock).numberCell(1, 0, "MyDecimal", new BigDecimal("8.2"));
			verify(handlerMock).numberCell(1, 1, "Integer", new BigDecimal("820"));
			verify(handlerMock).numberCell(1, 2, "MyCalculatedInteger", new BigDecimal("820"));

		} catch (InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}

	}
	

	@Test
	public void test_formattedNumbers() {
		
		try (OPCPackage pkg = OPCPackage.open(getClass().getResourceAsStream("/test-valid-formatted_numbers.xlsx"))) {
			
			XSSFWorkbook wb = new XSSFWorkbook(pkg);
			XSSFSheet sheet = wb.getSheetAt(0);
			
			XlsxSheetParser parser = new XlsxSheetParser();
			
			XlsxSheetEventHandler handlerMock = mock(XlsxSheetEventHandler.class);
			
			parser.parse(sheet, handlerMock);
			
			verify(handlerMock).startRow(1);
			
			verify(handlerMock).numberCell(1, 0, "MyFormattedDecimal", new BigDecimal("123456.78900"));
			verify(handlerMock).numberCell(1, 1, "MyFormattedInteger", new BigDecimal("987654321"));
			
		} catch (InvalidFormatException | IOException e) {
			throw new RuntimeException(e);
		}
		
	}

}
