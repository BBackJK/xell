package bback.module.xell.helper;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class ExcelSheetInfo {
    private final Sheet sheet;
    private final List<CellRangeAddress> mergeInfoList;
    private final List<Row> rowList;
    private final int maxRowIndex;
    private final int maxColumnIndex;

    public ExcelSheetInfo(Sheet sheet, List<CellRangeAddress> mergeInfoList, List<Row> rowList, int maxRowIndex, int maxColumnIndex) {
        this.sheet = sheet;
        this.mergeInfoList = mergeInfoList;
        this.rowList = rowList;
        this.maxRowIndex = maxRowIndex;
        this.maxColumnIndex = maxColumnIndex;
    }

    public Sheet getSheet() {
        return this.sheet;
    }

    public List<CellRangeAddress> getMergeInfoList() {
        return this.mergeInfoList;
    }

    public List<Row> getRowList() {
        return this.rowList;
    }

    public int getMaxRowIndex() {
        return this.maxRowIndex;
    }

    public int getMaxColumnIndex() {
        return this.maxColumnIndex;
    }

}
