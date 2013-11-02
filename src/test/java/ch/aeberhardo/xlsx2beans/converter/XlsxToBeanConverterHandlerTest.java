package ch.aeberhardo.xlsx2beans.converter;

import static org.junit.Assert.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.junit.Before;
import org.junit.Test;

import ch.aeberhardo.xlsx2beans.converter.beans.InvalidStringToNumberMappingBean;
import ch.aeberhardo.xlsx2beans.converter.beans.TestBean1;

public class XlsxToBeanConverterHandlerTest {

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void test_createRows() {

		XlsxToBeanConverterHandler<TestBean1> converterHandler = new XlsxToBeanConverterHandler<>(TestBean1.class);

		converterHandler.startRow(1);
		converterHandler.endRow(1);
		assertEquals(1, converterHandler.getBeans().size());

		converterHandler.startRow(2);
		converterHandler.endRow(2);
		assertEquals(2, converterHandler.getBeans().size());

	}

	@Test
	public void test_createBeans() throws ParseException {

		XlsxToBeanConverterHandler<TestBean1> converterHandler = new XlsxToBeanConverterHandler<>(TestBean1.class);

		converterHandler.startRow(1);
		converterHandler.stringCell(1, 0, "MyString1", "Test string 1");
		converterHandler.stringCell(1, 1, "MyString2", "Test string 2");
		converterHandler.numberCell(1, 2, "MyInteger", 123.0d);
		converterHandler.numberCell(1, 3, "MyDouble", 7.89d);
		converterHandler.dateCell(1, 4, "MyDate", new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").parse("12.01.2013 14:16:23"));
		converterHandler.endRow(1);

		assertEquals(1, converterHandler.getBeans().size());
		assertEquals("Test string 1", converterHandler.getBeans().get(0).getMyString1());
		assertEquals("Test string 2", converterHandler.getBeans().get(0).getMyString2());
		assertEquals(Integer.valueOf(123), converterHandler.getBeans().get(0).getMyInteger());
		assertEquals(Double.valueOf(7.89d), converterHandler.getBeans().get(0).getMyDouble());
		assertEquals(new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").parse("12.01.2013 14:16:23"), converterHandler.getBeans().get(0).getMyDate());

		converterHandler.startRow(2);
		converterHandler.stringCell(2, 0, "MyString1", "Test string 3");
		converterHandler.stringCell(2, 1, "MyString2", "Test string 4");
		converterHandler.numberCell(2, 2, "MyInteger", 567.0d);
		converterHandler.numberCell(2, 3, "MyDouble", 10.11d);
		converterHandler.dateCell(2, 4, "MyDate", new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").parse("20.06.2011 23:50:33"));
		converterHandler.endRow(2);

		assertEquals(2, converterHandler.getBeans().size());
		assertEquals("Test string 3", converterHandler.getBeans().get(1).getMyString1());
		assertEquals("Test string 4", converterHandler.getBeans().get(1).getMyString2());
		assertEquals(Integer.valueOf(567), converterHandler.getBeans().get(1).getMyInteger());
		assertEquals(Double.valueOf(10.11d), converterHandler.getBeans().get(1).getMyDouble());
		assertEquals(new SimpleDateFormat("dd.MM.yyyy HH:mm:ss").parse("20.06.2011 23:50:33"), converterHandler.getBeans().get(1).getMyDate());
	}

	@Test
	public void test_invalidStringToNumberMapping() {

		XlsxToBeanConverterHandler<InvalidStringToNumberMappingBean> converterHandler = new XlsxToBeanConverterHandler<>(InvalidStringToNumberMappingBean.class);

		converterHandler.startRow(1);

		try {

			converterHandler.stringCell(1, 0, "MyString1", "Test string 1");

			fail("Expected exception was not thrown!");

		} catch (XlsxBeanConverterException e) {
			assertTrue(e.getMessage().startsWith(
					"Error while mapping cell (rowNum=1, colIndex=0, colName=MyString1, cellValue=Test string 1): argument type mismatch"));
		}

	}

}
