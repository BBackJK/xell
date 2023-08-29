package bback.module.xell.util;

import java.lang.reflect.Field;

public final class ExcelStringUtils {

    private ExcelStringUtils() {
        throw new UnsupportedOperationException("This is a utility class and cannot be instantiated");
    }
    public static final String EMPTY = "";

    public static String toPascal(String value) {
        return value == null || value.isEmpty()
                ? EMPTY
                : value.replaceFirst(value.substring(0, 1), value.substring(0, 1).toUpperCase());
    }
    
    public static String toGetterByField(Field f) {
        return f == null ? EMPTY : toGetterByField(f.getName());
    }

    public static String toGetterByField(String fieldName) {
        return fieldName == null || fieldName.isEmpty()
                ? EMPTY
                : "get" + toPascal(fieldName);
    }

    public static String toSetterByField(Field f) {
        return f == null ? EMPTY : toSetterByField(f.getName());
    }

    public static String toSetterByField(String fieldName) {
        return fieldName == null || fieldName.isEmpty()
                ? EMPTY
                : "set" + toPascal(fieldName);
    }
}
