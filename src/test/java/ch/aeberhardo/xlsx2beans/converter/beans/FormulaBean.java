package ch.aeberhardo.xlsx2beans.converter.beans;

import java.util.Date;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class FormulaBean {

	private Integer m_myInteger;
	private Integer m_myFormulaInteger;

	private Date m_myFormulaTimestamp;

	public Integer getMyInteger() {
		return m_myInteger;
	}

	@XlsxColumnName("MyInteger")
	public void setMyInteger(Integer myInteger) {
		m_myInteger = myInteger;
	}

	public Integer getMyFormulaInteger() {
		return m_myFormulaInteger;
	}

	@XlsxColumnName("MyFormulaInteger")
	public void setMyFormulaInteger(Integer myFormulaInteger) {
		m_myFormulaInteger = myFormulaInteger;
	}

	public Date getMyFormulaTimestamp() {
		return m_myFormulaTimestamp;
	}

	@XlsxColumnName("MyFormulaTimestamp")
	public void setMyFormulaTimestamp(Date myFormulaTimestamp) {
		m_myFormulaTimestamp = myFormulaTimestamp;
	}
}
