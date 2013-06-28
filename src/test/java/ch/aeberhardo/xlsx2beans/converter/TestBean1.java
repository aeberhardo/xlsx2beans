package ch.aeberhardo.xlsx2beans.converter;

import java.util.Date;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class TestBean1 {

	private String m_myString1;
	private String m_myString2;
	private Integer m_myInteger;
	private Double m_myDouble;
	private Date m_myDate;

	public String getMyString1() {
		return m_myString1;
	}

	@XlsxColumnName("MyString1")
	public void setMyString1(String myString1) {
		m_myString1 = myString1;
	}

	public String getMyString2() {
		return m_myString2;
	}

	@XlsxColumnName("MyString2")
	public void setMyString2(String myString2) {
		m_myString2 = myString2;
	}

	public Integer getMyInteger() {
		return m_myInteger;
	}

	@XlsxColumnName("MyInteger")
	public void setMyInteger(Integer myInteger) {
		m_myInteger = myInteger;
	}

	public Double getMyDouble() {
		return m_myDouble;
	}

	@XlsxColumnName("MyDouble")
	public void setMyDouble(Double myDouble) {
		m_myDouble = myDouble;
	}

	public Date getMyDate() {
		return m_myDate;
	}

	@XlsxColumnName("MyDate")
	public void setMyDate(Date myDate) {
		m_myDate = myDate;
	}

}
