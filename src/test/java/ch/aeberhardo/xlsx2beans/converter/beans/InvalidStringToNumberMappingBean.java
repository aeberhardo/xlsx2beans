package ch.aeberhardo.xlsx2beans.converter.beans;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class InvalidStringToNumberMappingBean {

	@XlsxColumnName("MyString1")
	public void setMyInteger(Integer myString) {
		// nothing
	}

}
