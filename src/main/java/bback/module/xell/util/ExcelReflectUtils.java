package bback.module.xell.util;

import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;
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

@UtilityClass
@Slf4j
public class ExcelReflectUtils {

    public List<Method> filterFieldGetterMethods(@NonNull Class<?> clazz) {
        List<Method> result = new ArrayList<>();
        List<Field> localFields = filterLocalFields(clazz);
        int fieldCount = localFields.size();
        for (int i=0; i<fieldCount;i++) {
            try {
                Method m = clazz.getMethod(ExcelStringUtils.toGetterByField(localFields.get(i)));
                result.add(m);
            } catch (NoSuchMethodException ex) {
                log.error(ex.getMessage());
            }
        }
        return result;
    }

    public List<Field> filterLocalFields(@NonNull Class<?> clazz) {
        Objects.requireNonNull(clazz);
        return Arrays.stream(clazz.getDeclaredFields()).filter(f -> {
            int mod = f.getModifiers();
            return !Modifier.isFinal(mod) && !Modifier.isStatic(mod);
        }).collect(Collectors.toList());
    }

    public List<Field> filterFieldByExcelAnnotation(Class<?> clazz, Class<? extends Annotation> annotation) {
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
            log.info(e.getMessage());
            throw new RuntimeException(e);
        }
        return fieldList;
    }

    public List<Method> filterMethodByExcelAnnotation(Class<?> clazz, Class<? extends Annotation> annotation) {
        List<Method> methodList = new ArrayList<>();
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
            log.info(e.getMessage());
            throw new RuntimeException(e);
        }
        return methodList;
    }

    @Nullable
    public Object invokeTargetObject(MethodInvoker alreadySetTargetMethodInvoker, Object targetObject) {
        try {
            alreadySetTargetMethodInvoker.setTargetObject(targetObject);
            alreadySetTargetMethodInvoker.prepare();
            return alreadySetTargetMethodInvoker.invoke();
        } catch (ClassNotFoundException | NoSuchMethodException | InvocationTargetException | IllegalAccessException e) {
            log.warn(e.getMessage());
            // Skip..
            return null;
        }
    }

    public void invokeTargetData(MethodInvoker alreadySetTargetMethodInvoker, Object targetObject, Object... args) {
        if (alreadySetTargetMethodInvoker == null) return;
        try {
            alreadySetTargetMethodInvoker.setArguments(args);
            alreadySetTargetMethodInvoker.setTargetObject(targetObject);
            alreadySetTargetMethodInvoker.prepare();
            alreadySetTargetMethodInvoker.invoke();
        } catch (ClassNotFoundException | InvocationTargetException | NoSuchMethodException | IllegalAccessException e) {
            // skip..
            log.warn(e.getMessage());
            log.warn("{} 에 {} 메소드가 유효하지 않습니다.", targetObject, alreadySetTargetMethodInvoker.getTargetMethod());
        }
    }
}
