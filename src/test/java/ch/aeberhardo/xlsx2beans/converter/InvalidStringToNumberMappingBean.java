package ch.aeberhardo.xlsx2beans.converter;

public class InvalidStringToNumberMappingBean {

	@XlsxColumnName("MyString1")
	public void setMyInteger(Integer myString) {
		// nothing
	}

}
