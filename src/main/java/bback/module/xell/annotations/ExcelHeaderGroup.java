package bback.module.xell.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(value = RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelHeaderGroup {
    ExcelHeader[] value() default {};
    String[] merge() default {};
}
