package ch.aeberhardo.xlsx2beans.parser;

public class XlsxParserException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public XlsxParserException() {
		super();
	}

	public XlsxParserException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public XlsxParserException(String message, Throwable cause) {
		super(message, cause);
	}

	public XlsxParserException(String message) {
		super(message);
	}

	public XlsxParserException(Throwable cause) {
		super(cause);
	}

}
