package bback.module.xell.util;

import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.net.URL;
import java.util.*;

@UtilityClass
@Slf4j
public class ExcelClassUtils {
    public Set<Class<?>> scanClassByAnnotation(String packageName, Class<? extends Annotation> annotation) throws IOException, ClassNotFoundException {
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        String path = packageName.replace('.', '/');

        Set<Class<?>> classes = new LinkedHashSet<>();

        List<File> files = new ArrayList<>();
        Enumeration<URL> resources = classLoader.getResources(path);
        while (resources.hasMoreElements()) {
            URL resource = resources.nextElement();
            files.add(new File(resource.getFile()));
        }
        for (File file : files) {
            if (file.isDirectory()) {
                classes.addAll(findClasses(file, packageName, annotation));
            }
        }

        return classes;
    }

    public boolean isVoid(Class<?> classType) {
        return void.class.equals(classType) || Void.class.equals(classType);
    }

    private Set<Class<?>> findClasses(File directory, String packageName, Class<? extends Annotation> annotation) throws ClassNotFoundException {
        Set<Class<?>> classes = new LinkedHashSet<>();
        if (!directory.exists()) {
            return classes;
        }

        File[] files = directory.listFiles();
        if ( files != null ) {
            for (File file : files) {
                if (file.isDirectory()) {
                    classes.addAll(findClasses(file, packageName + "." + file.getName(), annotation));
                } else if (file.getName().endsWith(".class")) {
                    String className = packageName + '.' + file.getName().substring(0, file.getName().length() - 6);
                    ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
                    Class<?> clazz = Class.forName(className, false, classLoader);
                    Annotation anno = clazz.getAnnotation(annotation);
                    if ( anno != null ) {
                        classes.add(clazz);
                    }
                }
            }
        }

        return classes;
    }
}
