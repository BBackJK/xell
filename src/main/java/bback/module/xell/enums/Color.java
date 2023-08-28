package bback.module.xell.enums;

import org.apache.poi.hssf.util.HSSFColor;

public enum Color {

    BLACK(HSSFColor.HSSFColorPredefined.BLACK)
    , BROWN(HSSFColor.HSSFColorPredefined.BROWN)
    , OLIVE_GREEN(HSSFColor.HSSFColorPredefined.OLIVE_GREEN)
    , DARK_GREEN(HSSFColor.HSSFColorPredefined.DARK_GREEN)
    , DARK_TEAL(HSSFColor.HSSFColorPredefined.DARK_TEAL)
    , DARK_BLUE(HSSFColor.HSSFColorPredefined.DARK_BLUE)
    , INDIGO(HSSFColor.HSSFColorPredefined.INDIGO)
    , GREY_80_PERCENT(HSSFColor.HSSFColorPredefined.GREY_80_PERCENT)
    , ORANGE(HSSFColor.HSSFColorPredefined.ORANGE)
    , DARK_YELLOW(HSSFColor.HSSFColorPredefined.DARK_YELLOW)
    , GREEN(HSSFColor.HSSFColorPredefined.GREEN)
    , TEAL(HSSFColor.HSSFColorPredefined.TEAL)
    , BLUE(HSSFColor.HSSFColorPredefined.BLUE)
    , BLUE_GREY(HSSFColor.HSSFColorPredefined.BLUE_GREY)
    , GREY_50(HSSFColor.HSSFColorPredefined.GREY_50_PERCENT)
    , RED(HSSFColor.HSSFColorPredefined.RED)
    , LIGHT_ORANGE(HSSFColor.HSSFColorPredefined.LIGHT_ORANGE)
    , LIME(HSSFColor.HSSFColorPredefined.LIME)
    , SEA_GREEN(HSSFColor.HSSFColorPredefined.SEA_GREEN)
    , AQUA(HSSFColor.HSSFColorPredefined.AQUA)
    , LIGHT_BLUE(HSSFColor.HSSFColorPredefined.LIGHT_BLUE)
    , VIOLET(HSSFColor.HSSFColorPredefined.VIOLET)
    , GREY_40(HSSFColor.HSSFColorPredefined.GREY_40_PERCENT)
    , PINK(HSSFColor.HSSFColorPredefined.PINK)
    , GOLD(HSSFColor.HSSFColorPredefined.GOLD)
    , YELLOW(HSSFColor.HSSFColorPredefined.YELLOW)
    , BRIGHT_GREEN(HSSFColor.HSSFColorPredefined.BRIGHT_GREEN)
    , TURQUOISE(HSSFColor.HSSFColorPredefined.TURQUOISE)
    , DARK_RED(HSSFColor.HSSFColorPredefined.DARK_RED)
    , SKY_BLUE(HSSFColor.HSSFColorPredefined.SKY_BLUE)
    , PLUM(HSSFColor.HSSFColorPredefined.PLUM)
    , GREY_25(HSSFColor.HSSFColorPredefined.GREY_25_PERCENT)
    , ROSE(HSSFColor.HSSFColorPredefined.ROSE)
    , LIGHT_YELLOW(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW)
    , LIGHT_GREEN(HSSFColor.HSSFColorPredefined.LIGHT_GREEN)
    , LIGHT_TURQUOISE(HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE)
    , PALE_BLUE(HSSFColor.HSSFColorPredefined.PALE_BLUE)
    , LAVENDER(HSSFColor.HSSFColorPredefined.LAVENDER)
    , WHITE(HSSFColor.HSSFColorPredefined.WHITE)
    , CORNFLOWER_BLUE(HSSFColor.HSSFColorPredefined.CORNFLOWER_BLUE)
    , LEMON_CHIFFON(HSSFColor.HSSFColorPredefined.LEMON_CHIFFON)
    , MAROON(HSSFColor.HSSFColorPredefined.MAROON)
    , ORCHID(HSSFColor.HSSFColorPredefined.ORCHID)
    , CORAL(HSSFColor.HSSFColorPredefined.CORAL)
    , ROYAL_BLUE(HSSFColor.HSSFColorPredefined.ROYAL_BLUE)
    , LIGHT_CORNFLOWER_BLUE(HSSFColor.HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE)
    , TAN(HSSFColor.HSSFColorPredefined.TAN)
    ;
    private final HSSFColor.HSSFColorPredefined color;

    Color(HSSFColor.HSSFColorPredefined color) {
        this.color = color;
    }

    public short get() {
        return this.color.getIndex();
    }

}
