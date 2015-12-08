package ch.aeberhardo.xlsx2beans.converter;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class XlsxColumnNameToBeanMapper {

    public void setString(Object targetObject, String mappedColumnName, String value) {
        setObject(targetObject, mappedColumnName, value);
    }

    public void setDate(Object targetObject, String mappedColumnName, Date value) {
        setObject(targetObject, mappedColumnName, value);
    }

    public void setBoolean(Object targetObject, String mappedColumnName, Boolean value) {
        setObject(targetObject, mappedColumnName, value);
    }

    public void setNumber(Object targetObject, String mappedColumnName, BigDecimal value) {

        if (value == null) {
            return;
        }

        List<Method> methods = findMethodsByMappedColumnName(targetObject, mappedColumnName);

        for (Method method : methods) {
            Class<?> parameterType = method.getParameterTypes()[0];

            Object targetValue = null;

            if (parameterType == Integer.class) {
                targetValue = value.intValueExact();
            } else if (parameterType == Double.class) {
                targetValue = value.doubleValue();
            } else if (parameterType == Float.class) {
                targetValue = value.floatValue();
            } else if (parameterType == Short.class) {
                targetValue = value.shortValueExact();
            } else if (parameterType == Byte.class) {
                targetValue = value.byteValueExact();
            } else if (parameterType == Long.class) {
                targetValue = value.longValueExact();
            } else if (parameterType == BigDecimal.class) {
                targetValue = value;
            } else if (parameterType == String.class) {
                targetValue = value.toString();
            } else {
                throw new XlsxColumnNameToBeanMapperException("argument type mismatch! Mapping target expected " + parameterType.getSimpleName()
                        + " but got Number.");
            }

            invokeSetter(targetObject, method, targetValue);
        }
    }

    private void setObject(Object targetObject, String mappedColumnName, Object value) {

        if (value == null) {
            return;
        }

        List<Method> methods = findMethodsByMappedColumnName(targetObject, mappedColumnName);

        for (Method method : methods) {
            invokeSetter(targetObject, method, value);
        }
    }

    private void invokeSetter(Object targetObject, Method method, Object value) {

        validateSetterName(method);

        try {
            method.invoke(targetObject, value);
        } catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
            throw new XlsxColumnNameToBeanMapperException(e.getMessage(), e);
        }
    }

    private void validateSetterName(Method method) {
        if (!method.getName().startsWith("set")) {
            throw new XlsxColumnNameToBeanMapperException("Name of setter method must start with 'set' but was '" + method.getName() + "'.");
        }
    }

    private List<Method> findMethodsByMappedColumnName(Object targetObject, String mappedColumnName) {

        List<Method> methods = new ArrayList<>();

        for (Method method : targetObject.getClass().getDeclaredMethods()) {

            if (method.isAnnotationPresent(XlsxColumnName.class)) {

                XlsxColumnName annotation = method.getAnnotation(XlsxColumnName.class);
                String value = annotation.value();

                if (mappedColumnName.equals(value)) {
                    methods.add(method);
                }
            }
        }

        //look for annotation on its superclass if it exists
        if (targetObject.getClass().getSuperclass() != null) {
            for (Method method : targetObject.getClass().getSuperclass().getDeclaredMethods()) {

                if (method.isAnnotationPresent(XlsxColumnName.class)) {

                    XlsxColumnName annotation = method.getAnnotation(XlsxColumnName.class);
                    String value = annotation.value();

                    if (mappedColumnName.equals(value)) {
                        methods.add(method);
                    }
                }
            }
        }

        return methods;
    }

}
