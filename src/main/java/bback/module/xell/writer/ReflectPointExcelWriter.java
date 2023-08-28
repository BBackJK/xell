package bback.module.xell.writer;

import bback.module.xell.annotations.CellConfig;
import bback.module.xell.annotations.ExcelPointer;
import bback.module.xell.annotations.FontConfig;
import bback.module.xell.helper.ExcelMethodInvokerHelper;
import bback.module.xell.util.ExcelReflectUtils;
import bback.module.xell.util.ExcelStringUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.MethodInvoker;

import java.util.regex.Pattern;

@Slf4j
public class ReflectPointExcelWriter<T> extends AbstractPointExcelWriter<T> {


    public ReflectPointExcelWriter(Class<T> classType) {
        this(classType, null);
    }

    public ReflectPointExcelWriter(Class<T> classType, T data) {
        super(classType, data);
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
            if (super.excelMethodInvokerHelperMap.isEmpty() && getPattern().matcher(sourceCellValue).find()) {
                String propertyValue = sourceCellValue.replaceAll("[\\W]", "");
                MethodInvoker invoker = new MethodInvoker();
                try {
                    invoker.setTargetObject(this.data);
                    invoker.setTargetMethod(ExcelStringUtils.toGetterByField(propertyValue));
                    invoker.prepare();
                    Object cellValue = invoker.invoke();
                    target.setCellValue(cellValue == null ? "" : String.valueOf(cellValue));
                } catch (Exception e) {
                    log.warn(e.getMessage());
                    target.setCellValue("");
                }
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

        } else {
            String cellAddr = target.getAddress().formatAsString();
            ExcelMethodInvokerHelper<ExcelPointer> excelPointHelper = this.excelMethodInvokerHelperMap.get(cellAddr);
            if (excelPointHelper != null) {
                // 값 바인딩
                MethodInvoker invoker = excelPointHelper.getInvoker();
                Object cellValue = ExcelReflectUtils.invokeTargetObject(invoker, this.data);
                target.setCellValue(cellValue == null ? "" : String.valueOf(cellValue));

                excelPointHelper.getAnnotation().ifPresent(excelPointer -> {
                    // font style
                    Font copiedFont = targetWorkbook.createFont();
                    FontConfig pointerFont = excelPointer.font();
                    copiedFont.setColor(pointerFont.color().get());
                    copiedFont.setBold(pointerFont.isBold());

                    // cell style
                    CellStyle copiedCellStyle = targetWorkbook.createCellStyle();
                    CellConfig pointerCell = excelPointer.cell();
                    copiedCellStyle.setFillForegroundColor(pointerCell.color().get());
                    copiedCellStyle.setVerticalAlignment(pointerCell.va());
                    copiedCellStyle.setAlignment(pointerCell.ha());
                    copiedCellStyle.setFillPattern(pointerCell.fill());
                    copiedCellStyle.setFont(copiedFont);

                    // apply cell style
                    target.setCellStyle(copiedCellStyle);
                });
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
