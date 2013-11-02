package ch.aeberhardo.xlsx2beans.converter.beans;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class NumberAsTextBean {

	private String m_myString;

	public String getMyString() {
		return m_myString;
	}

	@XlsxColumnName("MyString")
	public void setMyString(String myString) {
		m_myString = myString;
	}

}
