package ch.aeberhardo.xlsx2beans.converter;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import ch.aeberhardo.xlsx2beans.parser.XlsxSheetEventHandler;

public class XlsxToBeanConverterHandler<T> implements XlsxSheetEventHandler {

	private final Class<T> m_beanType;

	private final List<T> m_beans = new ArrayList<>();

	private final XlsxColumnNameToBeanMapper m_mapper = new XlsxColumnNameToBeanMapper();

	public XlsxToBeanConverterHandler(Class<T> beanType) {
		m_beanType = beanType;
	}

	@Override
	public void startRow(int rowNum) {
		T bean = createNewBean();
		m_beans.add(bean);
	}

	@Override
	public void endRow(int rowNum) {
		// Nothing
	}

	@Override
	public void stringCell(int rowNum, int colIndex, String colName, String cellValue) {
		T currentBean = getCurrentBean();
		try {
			m_mapper.setString(currentBean, colName, cellValue);
		} catch (XlsxColumnNameToBeanMapperException e) {
			throw beanConverterException(rowNum, colIndex, colName, cellValue, e);
		}
	}

	@Override
	public void numberCell(int rowNum, int colIndex, String colName, Number cellValue) {
		T currentBean = getCurrentBean();
		try {
			m_mapper.setNumber(currentBean, colName, cellValue);
		} catch (XlsxColumnNameToBeanMapperException e) {
			throw beanConverterException(rowNum, colIndex, colName, cellValue.toString(), e);
		}
	}

	@Override
	public void dateCell(int rowNum, int colIndex, String colName, Date cellValue) {
		T currentBean = getCurrentBean();
		try {
			m_mapper.setDate(currentBean, colName, cellValue);
		} catch (XlsxColumnNameToBeanMapperException e) {
			throw beanConverterException(rowNum, colIndex, colName, cellValue.toString(), e);
		}
	}

	private T createNewBean() {
		try {
			return m_beanType.newInstance();
		} catch (InstantiationException | IllegalAccessException e) {
			throw new XlsxBeanConverterException(e);
		}
	}

	private T getCurrentBean() {
		return m_beans.get(m_beans.size() - 1);
	}

	public List<T> getBeans() {
		return Collections.unmodifiableList(m_beans);
	}

	private XlsxBeanConverterException beanConverterException(int rowNum, int colIndex, String colName, String cellValueAsString,
			XlsxColumnNameToBeanMapperException cause) {
		String message = "Error while mapping cell (rowNum=" + rowNum + ", colIndex=" + colIndex + ", colName=" + colName + ", cellValue=" + cellValueAsString
				+ "): " + cause.getMessage();
		return new XlsxBeanConverterException(message, cause);
	}
}
