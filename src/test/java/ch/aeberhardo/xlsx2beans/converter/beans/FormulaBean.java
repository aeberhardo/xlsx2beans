package ch.aeberhardo.xlsx2beans.converter.beans;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class FormulaBean {

	private Integer m_myInteger;
	private Integer m_myFormulaInteger;

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

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("FormulaBean [m_myInteger=");
		builder.append(m_myInteger);
		builder.append(", m_myFormulaInteger=");
		builder.append(m_myFormulaInteger);
		builder.append("]");
		return builder.toString();
	}

}
