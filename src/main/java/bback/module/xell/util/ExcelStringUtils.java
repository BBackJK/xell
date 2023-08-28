package bback.module.xell.util;

import lombok.experimental.UtilityClass;

import java.lang.reflect.Field;

@UtilityClass
public class ExcelStringUtils {
    public final String EMPTY = "";

    public String toPascal(String value) {
        return value == null || value.isEmpty()
                ? EMPTY
                : value.replaceFirst(value.substring(0, 1), value.substring(0, 1).toUpperCase());
    }
    
    public String toGetterByField(Field f) {
        return f == null ? EMPTY : toGetterByField(f.getName());
    }

    public String toGetterByField(String fieldName) {
        return fieldName == null || fieldName.isEmpty()
                ? EMPTY
                : "get" + toPascal(fieldName);
    }

    public String toSetterByField(Field f) {
        return f == null ? EMPTY : toSetterByField(f.getName());
    }

    public String toSetterByField(String fieldName) {
        return fieldName == null || fieldName.isEmpty()
                ? EMPTY
                : "set" + toPascal(fieldName);
    }
}
