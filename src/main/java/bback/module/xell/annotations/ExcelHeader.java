package bback.module.xell.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(value = RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelHeader {

    String[] value() default {"헤더1", "헤더2", "헤더3", "헤더4"};
    FontConfig font() default @FontConfig;
    CellConfig cell() default @CellConfig;
}
