package bback.module.xell.annotations;


import bback.module.xell.enums.Color;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(value = RetentionPolicy.RUNTIME)
@Target({})
public @interface FontConfig {
    Color color() default Color.BLACK;
    boolean isBold() default false;
}
