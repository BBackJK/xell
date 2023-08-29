package bback.module.xell.writer;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.util.Map;
import java.util.regex.Pattern;

@Slf4j
public class MapPointExcelWriter extends AbstractPointExcelWriter<Map> {


    public MapPointExcelWriter() {
        this(null);
    }

    public MapPointExcelWriter(Map<String, String> data) {
        super(Map.class, data);
    }

    @Override
    protected void setRow(Sheet sourceSheet, Sheet targetSheet, Row source, Row target) {
        target.setHeight(source.getHeight());
        target.setHeightInPoints(source.getHeightInPoints());
        target.setZeroHeight(source.getZeroHeight());

        CellStyle rowCellStyle = source.getRowStyle();
        if ( rowCellStyle != null ) {
            CellStyle targetCellStyle = targetSheet.getWorkbook().createCellStyle();
            targetCellStyle.cloneStyleFrom(rowCellStyle);
            target.setRowStyle(targetCellStyle);
        }
    }

    @Override
    protected void setCell(Sheet sourceSheet, Sheet targetSheet, Cell source, Cell target) {
        Workbook targetWorkbook = targetSheet.getWorkbook();

        if (source != null) {
            // 값 바인딩
            String sourceCellValue = source.getStringCellValue();
            if (getPattern().matcher(sourceCellValue).find()) {
                String propertyValue = sourceCellValue.replaceAll("[\\W]", "");
                String cellValue = this.data == null ? "" : (String) this.data.get(propertyValue);
                target.setCellValue(cellValue == null ? "" : cellValue);
            } else {
                target.setCellValue(sourceCellValue);
            }

            // cell style 바인딩
            CellStyle sourceCellStyle = source.getCellStyle();
            if (sourceCellStyle != null) {
                CellStyle targetCellStyle = targetWorkbook.createCellStyle();
                targetCellStyle.cloneStyleFrom(sourceCellStyle);
                target.setCellStyle(targetCellStyle);
            }
        }

        // width 바인딩
        int columnIndex = target.getColumnIndex();
        targetSheet.setColumnWidth(columnIndex, sourceSheet.getColumnWidth(columnIndex));
    }

    private Pattern getPattern() {
        return Pattern.compile(this.startPattern+"[\\s]?[a-z|0-9]+[\\s]?"+this.endPattern);
    }
}
