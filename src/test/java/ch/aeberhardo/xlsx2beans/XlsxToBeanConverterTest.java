package ch.aeberhardo.xlsx2beans;

import static org.junit.Assert.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;

import org.junit.Before;
import org.junit.Test;

import ch.aeberhardo.xlsx2beans.XlsxToBeanConverter;
import ch.aeberhardo.xlsx2beans.converter.TestBean1;

public class XlsxToBeanConverterTest {

	@Before
	public void setUp() throws Exception {
	}

	@Test
	public void test_convert() throws ParseException {

		XlsxToBeanConverter converter = new XlsxToBeanConverter();

		List<TestBean1> beans = converter.convert(getClass().getResourceAsStream("/test1.xlsx"), 0, TestBean1.class);

		assertEquals(2, beans.size());

		assertEquals("ABC", beans.get(0).getMyString1());
		assertEquals("This is my string 1", beans.get(0).getMyString2());
		assertEquals(Integer.valueOf(123), beans.get(0).getMyInteger());
		assertEquals(Double.valueOf(7.89), beans.get(0).getMyDouble());
		assertEquals(new SimpleDateFormat("dd.MM.yyyy hh:mm:ss").parse("12.01.2013 14:16:23"), beans.get(0).getMyDate());

		assertEquals("DEF", beans.get(1).getMyString1());
		assertEquals("This is my string 2", beans.get(1).getMyString2());
		assertEquals(Integer.valueOf(456), beans.get(1).getMyInteger());
		assertEquals(Double.valueOf(9.87), beans.get(1).getMyDouble());
		assertEquals(new SimpleDateFormat("dd.MM.yyyy hh:mm:ss").parse("20.06.2011 23:50:33"), beans.get(1).getMyDate());
	}

	@Test
	public void test_convertWithEmptyCells() throws ParseException {

		XlsxToBeanConverter converter = new XlsxToBeanConverter();

		List<TestBean1> beans = converter.convert(getClass().getResourceAsStream("/test-valid_empty_cells.xlsx"), 0, TestBean1.class);

		assertEquals(2, beans.size());

		assertEquals("ABC", beans.get(0).getMyString1());
		assertEquals(null, beans.get(0).getMyString2());
		assertEquals(Integer.valueOf(123), beans.get(0).getMyInteger());
		assertEquals(Double.valueOf(7.89), beans.get(0).getMyDouble());
		assertEquals(null, beans.get(0).getMyDate());

		assertEquals("DEF", beans.get(1).getMyString1());
		assertEquals("This is my string 2", beans.get(1).getMyString2());
		assertEquals(Integer.valueOf(456), beans.get(1).getMyInteger());
		assertEquals(null, beans.get(1).getMyDouble());
		assertEquals(new SimpleDateFormat("dd.MM.yyyy hh:mm:ss").parse("20.06.2011 23:50:33"), beans.get(1).getMyDate());
	}

}
