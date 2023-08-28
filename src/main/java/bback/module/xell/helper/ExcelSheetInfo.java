package bback.module.xell.helper;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

@Getter
@AllArgsConstructor
public class ExcelSheetInfo {
    private final Sheet sheet;
    private final List<CellRangeAddress> mergeInfoList;
    private final List<Row> rowList;
    private final int lastRowIndex;
    private final int lastColumnIndex;


}
