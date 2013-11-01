package ch.aeberhardo.xlsx2beans.converter.beans;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class DoublePrecisionBean {

	private Double m_myDecimal;
	private Integer m_myInteger;
	private Integer m_myCalculatedInteger;

	public Double getMyDecimal() {
		return m_myDecimal;
	}

	@XlsxColumnName("MyDecimal")
	public void setMyDecimal(Double myDecimal) {
		m_myDecimal = myDecimal;
	}

	public Integer getMyInteger() {
		return m_myInteger;
	}

	@XlsxColumnName("Integer")
	public void setMyInteger(Integer myInteger) {
		m_myInteger = myInteger;
	}

	public Integer getMyCalculatedInteger() {
		return m_myCalculatedInteger;
	}

	@XlsxColumnName("MyCalculatedInteger")
	public void setMyCalculatedInteger(Integer myCalculatedInteger) {
		m_myCalculatedInteger = myCalculatedInteger;
	}

}
