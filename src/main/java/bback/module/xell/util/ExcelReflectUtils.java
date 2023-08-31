package bback.module.xell.util;

import bback.module.xell.logger.Log;
import bback.module.xell.logger.LogFactory;
import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;
import org.springframework.util.MethodInvoker;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

public final class ExcelReflectUtils {

    private static final Log LOGGER = LogFactory.getLog(ExcelReflectUtils.class);

    private ExcelReflectUtils() {
        throw new UnsupportedOperationException("This is a utility class and cannot be instantiated");
    }

    public static List<Method> filterFieldGetterMethods(@NonNull Class<?> clazz) {
        List<Method> result = new ArrayList<>();
        List<Field> localFields = filterLocalFields(clazz);
        int fieldCount = localFields.size();
        for (int i=0; i<fieldCount;i++) {
            try {
                Method m = clazz.getMethod(ExcelStringUtils.toGetterByField(localFields.get(i)));
                result.add(m);
            } catch (NoSuchMethodException ex) {
                LOGGER.error(ex.getMessage());
            }
        }
        return result;
    }

    public static List<Field> filterLocalFields(@NonNull Class<?> clazz) {
        Objects.requireNonNull(clazz);
        return Arrays.stream(clazz.getDeclaredFields()).filter(f -> {
            int mod = f.getModifiers();
            return !Modifier.isFinal(mod) && !Modifier.isStatic(mod);
        }).collect(Collectors.toList());
    }

    public static List<Field> filterFieldByExcelAnnotation(Class<?> clazz, Class<? extends Annotation> annotation) {
        List<Field> fieldList = new ArrayList<>();
        if ( clazz == null ) {
            return fieldList;
        }

        try {
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                Annotation annotated = field.getAnnotation(annotation);
                if ( annotated != null ) {
                    fieldList.add(field);
                }
            }
        } catch (SecurityException e) {
            LOGGER.warn(e.getMessage());
        }
        return fieldList;
    }

    public static List<Method> filterMethodByExcelAnnotation(Class<?> clazz, Class<? extends Annotation> annotation) {
        List<Method> methodList = new ArrayList<>();
        if ( clazz == null ) {
            return methodList;
        }
        try {

            Method[] methods = clazz.getDeclaredMethods();
            int methodLength = methods.length;

            for (int i=0;i<methodLength;i++) {
                Method method = methods[i];
                if ( method.getAnnotation(annotation) != null ) {
                    methodList.add(method);
                }
            }

        } catch (SecurityException e) {
            LOGGER.warn(e.getMessage());
        }
        return methodList;
    }

    @Nullable
    public static Object invokeTargetObject(MethodInvoker alreadySetTargetMethodInvoker, Object targetObject) {
        try {
            alreadySetTargetMethodInvoker.setTargetObject(targetObject);
            alreadySetTargetMethodInvoker.prepare();
            return alreadySetTargetMethodInvoker.invoke();
        } catch (ClassNotFoundException | NoSuchMethodException | InvocationTargetException | IllegalAccessException e) {
            LOGGER.warn(e.getMessage());
            // Skip..
            return null;
        }
    }

    public static void invokeTargetData(MethodInvoker alreadySetTargetMethodInvoker, Object targetObject, Object... args) {
        if (alreadySetTargetMethodInvoker == null) return;
        try {
            alreadySetTargetMethodInvoker.setArguments(args);
            alreadySetTargetMethodInvoker.setTargetObject(targetObject);
            alreadySetTargetMethodInvoker.prepare();
            alreadySetTargetMethodInvoker.invoke();
        } catch (ClassNotFoundException | InvocationTargetException | NoSuchMethodException | IllegalAccessException e) {
            // skip..
            LOGGER.warn(e.getMessage());
            LOGGER.warn(targetObject + " 에 " + alreadySetTargetMethodInvoker.getTargetMethod() + " 메소드가 유효하지 않습니다.");
        }
    }
}
