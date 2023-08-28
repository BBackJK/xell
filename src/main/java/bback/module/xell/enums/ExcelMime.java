package bback.module.xell.enums;

public enum ExcelMime {
    XLSX("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    , XLS("application/vnd.ms-excel");

    private final String value;

    ExcelMime(String value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return this.value;
    }
}
