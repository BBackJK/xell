package bback.module.xell.writer;

import bback.module.xell.annotations.*;
import bback.module.xell.exceptions.ExcelWriteException;
import bback.module.xell.helper.ExcelMethodInvokerHelper;
import bback.module.xell.logger.Log;
import bback.module.xell.logger.LogFactory;
import bback.module.xell.util.ExcelListUtils;
import bback.module.xell.util.ExcelReflectUtils;
import bback.module.xell.util.ExcelStringUtils;
import bback.module.xell.util.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.core.annotation.AnnotationUtils;
import org.springframework.lang.NonNull;
import org.springframework.util.MethodInvoker;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class ReflectGridExcelWriter<T> extends AbstractGridExcelWriter<T> {

    private static final Log LOGGER = LogFactory.getLog(ReflectGridExcelWriter.class);

    public ReflectGridExcelWriter(Class<T> classType) {
        super(classType, Collections.emptyList());
    }

    public ReflectGridExcelWriter(Class<T> classType, List<T> dataList) {
        super(classType, dataList);
    }

    @Override
    protected void writeHeader(SXSSFWorkbook wb, SXSSFSheet sheet) throws ExcelWriteException {
        List<ExcelHeader> excelHeaders = new ArrayList<>();
        ExcelHeaderGroup excelHeaderGroup = AnnotationUtils.findAnnotation(super.classType, ExcelHeaderGroup.class);
        if ( excelHeaderGroup != null ) {
            excelHeaders.addAll(Arrays.asList(excelHeaderGroup.value()));
        } else {
            excelHeaders.add(AnnotationUtils.findAnnotation(super.classType, ExcelHeader.class));
        }

        if ( excelHeaders.isEmpty() ) {
            throw new ExcelWriteException("엑셀 데이터를 만들 수 없습니다.");
        }

        int headerRows = excelHeaders.size();
        for (int i=0;i<headerRows;i++) {
            ExcelHeader excelHeader = excelHeaders.get(i);

            String[] headers = excelHeader.value();
            CellConfig cell = excelHeader.cell();
            FontConfig font = excelHeader.font();

            // make font
            Font fontStyle = wb.createFont();
            fontStyle.setBold(font.isBold());
            fontStyle.setColor(font.color().get());

            // make cell style
            CellStyle cellStyle = wb.createCellStyle();
            cellStyle.setFillForegroundColor(cell.color().get());
            cellStyle.setFillPattern(cell.fill());
            cellStyle.setAlignment(cell.ha());
            cellStyle.setVerticalAlignment(cell.va());
            cellStyle.setFont(fontStyle);

            // body row 생성
            Row headerRow = sheet.createRow(i);
            int headerCellSize = headers.length;
            for(int j =0; j<headerCellSize; j++) {
                Cell headerCell = headerRow.createCell(j);
                headerCell.setCellStyle(cellStyle);
                headerCell.setCellValue(headers[j]);
            }
        }

        // merge
        if ( excelHeaderGroup != null ) {
            String[] mergeCells = excelHeaderGroup.merge();
            int mergeInfoSize = mergeCells.length;
            for (int z=0; z<mergeInfoSize; z++) {
                String mergeCellValue = mergeCells[z];
                try {
                    sheet.addMergedRegion(CellRangeAddress.valueOf(mergeCellValue));
                } catch (IllegalArgumentException e) {
                    LOGGER.error(" 병합 수식 [" + mergeCellValue + "] 는 유효하지 않은 값입니다. ", e);
                }
            }
        }
    }

    @Override
    protected void writeBody(SXSSFWorkbook wb, SXSSFSheet sheet) throws ExcelWriteException {
        if ( this.dataList == null || this.dataList.isEmpty() ) {
            return;
        }

        List<ExcelMethodInvokerHelper<ExcelBody>> excelBodyInvokerHelperList = this.getExcelMethodInvokerHelper(super.classType);
        Collections.sort(excelBodyInvokerHelperList);  // 정렬

        int lastHeaderRowIndex = sheet.getLastRowNum();
        Row lastHeaderRow = sheet.getRow(lastHeaderRowIndex);
        int cellCount = lastHeaderRow.getPhysicalNumberOfCells();   // body row 당 cell 사이즈 (마지막 Header 의 cell 수를 차용.)
        int dataCount = this.dataList.size();                       // body row 사이즈
        int rowNum = lastHeaderRowIndex + 1; // rowNum

        List<CellStyle> cellStyles = new ArrayList<>(cellCount);
        for (int i=0; i<dataCount ;i++) {
            Row row = sheet.createRow(rowNum++);
            for (int j=0; j<cellCount;j++) {
                Cell cell = row.createCell(j);

                ExcelMethodInvokerHelper<ExcelBody> invokerHelper = ExcelListUtils.getOrSupply(excelBodyInvokerHelperList, j);
                if (invokerHelper == null) continue;
                MethodInvoker invoker = invokerHelper.getInvoker();
                Object cellValue = ExcelReflectUtils.invokeTargetObject(invoker, dataList.get(i));
                // value 바인딩
                cell.setCellValue(cellValue == null ? "" : String.valueOf(cellValue));

                // cell style 바인딩
                CellStyle cellStyle = ExcelListUtils.getOrSupply(cellStyles, j, () -> {
                    CellStyle supplyCellStyle = wb.createCellStyle();
                    Font supplyCellFont = wb.createFont();

                    invokerHelper.getAnnotation().ifPresent(excelBody -> {
                        // apply excel body font config
                        FontConfig fontConfig = excelBody.font();
                        supplyCellFont.setBold(fontConfig.isBold());
                        supplyCellFont.setColor(fontConfig.color().get());

                        // apply excel body cell config
                        CellConfig cellConfig = excelBody.cell();
                        supplyCellStyle.setFillForegroundColor(cellConfig.color().get());
                        supplyCellStyle.setFillPattern(cellConfig.fill());
                        supplyCellStyle.setAlignment(cellConfig.ha());
                        supplyCellStyle.setVerticalAlignment(cellConfig.va());
                    });

                    supplyCellStyle.setFont(supplyCellFont);
                    return supplyCellStyle;
                });
                cell.setCellStyle(cellStyle);
            }

            if ( rowNum % ExcelUtils.WINDOW_SIZE == 0 ) {
                try {
                    sheet.flushRows(ExcelUtils.WINDOW_SIZE);
                } catch (IOException e) {
                    LOGGER.error("엑셀 파일을 만드는 중 Sheet 를 flush 시키는데 실패하였습니다.");
                    throw new ExcelWriteException();
                }
            }
        }
    }

    @NonNull
    private List<ExcelMethodInvokerHelper<ExcelBody>> getExcelMethodInvokerHelper(Class<?> clazz) {
        if (clazz == null) return Collections.emptyList();

        List<ExcelMethodInvokerHelper<ExcelBody>> result = new ArrayList<>();

        List<Field> excelBodyFieldList = ExcelReflectUtils.filterFieldByExcelAnnotation(clazz, ExcelBody.class);
        List<Method> excelBodyMethodList = ExcelReflectUtils.filterMethodByExcelAnnotation(clazz, ExcelBody.class);

        int excelBodyFieldCount = excelBodyFieldList.size();
        int excelBodyMethodCount = excelBodyMethodList.size();

        for (int i=0; i<excelBodyFieldCount;i++) {
            Field field = excelBodyFieldList.get(i);
            ExcelBody excelBody = field.getAnnotation(ExcelBody.class);
            MethodInvoker methodInvoker = new MethodInvoker();
            methodInvoker.setTargetMethod(ExcelStringUtils.toGetterByField(field));
            int value = excelBody.value();
            int order = excelBody.order();
            int sort = value == 0 ? order : value;
            result.add(
                    new ExcelMethodInvokerHelper<>(
                            methodInvoker, excelBody, sort
                    )
            );
        }

        for (int i=0; i<excelBodyMethodCount;i++) {
            Method method = excelBodyMethodList.get(i);
            ExcelBody excelBody = method.getAnnotation(ExcelBody.class);
            MethodInvoker methodInvoker = new MethodInvoker();
            methodInvoker.setTargetMethod(method.getName());
            int value = excelBody.value();
            int order = excelBody.order();
            int sort = value == 0 ? order : value;
            result.add(
                    new ExcelMethodInvokerHelper<>(
                            methodInvoker, excelBody, sort
                    )
            );
        }

        if ( result.isEmpty() ) { // @ExcelBody 어노테이션이 없다면..
            List<Method> getterMethods = ExcelReflectUtils.filterFieldGetterMethods(clazz);
            int methodCount = getterMethods.size();
            for (int i=0; i<methodCount;i++) {
                Method m = getterMethods.get(i);
                MethodInvoker methodInvoker = new MethodInvoker();
                methodInvoker.setTargetMethod(m.getName());
                result.add(
                        new ExcelMethodInvokerHelper<>(
                                methodInvoker, null, 0
                        )
                );
            }
        }

        return result;
    }
}
