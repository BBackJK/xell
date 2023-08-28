package bback.module.xell.annotations;


import bback.module.xell.enums.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(value = RetentionPolicy.RUNTIME)
@Target({})
public @interface CellConfig {

    Color color() default Color.WHITE;
    FillPatternType fill() default FillPatternType.SOLID_FOREGROUND;
    HorizontalAlignment ha() default HorizontalAlignment.CENTER;
    VerticalAlignment va() default VerticalAlignment.CENTER;
}
