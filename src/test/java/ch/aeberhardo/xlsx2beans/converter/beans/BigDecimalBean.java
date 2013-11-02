package ch.aeberhardo.xlsx2beans.converter.beans;

import java.math.BigDecimal;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class BigDecimalBean {

	private BigDecimal m_myBigDecimal;

	public BigDecimal getMyBigDecimal() {
		return m_myBigDecimal;
	}

	@XlsxColumnName("MyBigDecimal")
	public void setMyBigDecimal(BigDecimal myBigDecimal) {
		m_myBigDecimal = myBigDecimal;
	}


}
