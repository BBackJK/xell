package bback.module.xell.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(value = RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.METHOD})
public @interface ExcelPointer {

    String value() default "";     // A1, A2, ...
    boolean isFormula() default false; // 수식 여부
    FontConfig font() default @FontConfig;
    CellConfig cell() default @CellConfig;
}
