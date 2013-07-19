package ch.aeberhardo.xlsx2beans.converter.beans;

import java.util.Date;

import ch.aeberhardo.xlsx2beans.converter.XlsxColumnName;

public class DateBean {

	private Date m_timestamp;
	private Date m_date;
	private Date m_time;

	public Date getTimestamp() {
		return m_timestamp;
	}

	@XlsxColumnName("MyTimestamp")
	public void setTimestamp(Date timestamp) {
		m_timestamp = timestamp;
	}

	public Date getDate() {
		return m_date;
	}

	@XlsxColumnName("MyDate")
	public void setDate(Date date) {
		m_date = date;
	}

	public Date getTime() {
		return m_time;
	}

	@XlsxColumnName("MyTime")
	public void setTime(Date time) {
		m_time = time;
	}

}
