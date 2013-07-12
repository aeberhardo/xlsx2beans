package ch.aeberhardo.xlsx2beans.converter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 
 * Annotation used to provide mapping information.
 * It is used to annotate the setter method on a Java bean class that should receive data of a specified XLSX spread sheet column.
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
public @interface XlsxColumnName {

	/**
	 * The name of the XLSX spread sheet column that provides data for the setter method.
	 * The first row of the  spread sheet has to contain the column names.
	 */
	String value();

}
