package bback.module.xell.reader;

import bback.module.xell.annotations.ExcelTitle;
import bback.module.xell.exceptions.ExcelReadException;
import bback.module.xell.exceptions.ExcelWriteException;
import bback.module.xell.util.ExcelReflectUtils;
import bback.module.xell.util.ExcelStringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.lang.Nullable;
import org.springframework.util.MethodInvoker;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class ReflectGridExcelReader<T> extends AbstractGridExcelReader<T, MethodInvoker> {

    private static final Logger LOGGER = LoggerFactory.getLogger(ReflectGridExcelReader.class);

    public ReflectGridExcelReader(Class<T> classType) {
        super(classType);
    }

    public ReflectGridExcelReader(Class<T> classType, String filename) throws IOException, ExcelReadException {
        this(classType, Files.newInputStream(Paths.get(filename)));
    }

    public ReflectGridExcelReader(Class<T> classType, InputStream is) throws ExcelReadException {
        super(classType, is);
    }

    @Override
    protected List<MethodInvoker> makeHeaderValueList(Sheet sheet, int headerSkipCount) {
        int totalRowCount = sheet.getPhysicalNumberOfRows();
        if (totalRowCount < 1 || totalRowCount < headerSkipCount) {
            throw new ExcelWriteException();
        }

        List<Field> excelTitleFields = ExcelReflectUtils.filterFieldByExcelAnnotation(classType, ExcelTitle.class);
        List<MethodInvoker> invokers = new ArrayList<>();

        boolean hasExcelTitleField = true;
        List<Field> targetFields = excelTitleFields;
        if ( excelTitleFields.isEmpty() ) {
            targetFields = ExcelReflectUtils.filterLocalFields(this.classType);
            hasExcelTitleField = false;
        }

        int lastHeaderRowIndex = headerSkipCount -1;

        Row row = sheet.getRow(lastHeaderRowIndex);

        int totalColumns = row.getPhysicalNumberOfCells();
        for (int j=0; j<totalColumns; j++) {    // row 의 cell 반복
            Cell cell = row.getCell(j);
            if (cell == null) continue;

            for (Field f : targetFields) {
                MethodInvoker invoker = new MethodInvoker();
                if ( hasExcelTitleField ) {
                    ExcelTitle excelTitle = f.getAnnotation(ExcelTitle.class);
                    String cellValue = cell.getStringCellValue();
                    if ( excelTitle.value().equals(cellValue) ) {   // ExcelTitle 이 있는 경우는 cell Header 값과 ExcelTitle value 값이 동일해야 적용
                        invoker.setTargetMethod(ExcelStringUtils.toSetterByField(f));
                        invokers.add(invoker);
                    }
                } else {
                    invoker.setTargetMethod(ExcelStringUtils.toSetterByField(f));
                    invokers.add(invoker);
                }
            }
        }
        return invokers;
    }

    @Override
    protected T createTarget() throws ExcelReadException {
        try {
            Constructor<T> constructor = super.classType.getConstructor();
            return constructor.newInstance();
        } catch (Exception e) {
            LOGGER.error(String.format("%s 클래스의 생성자에 접근할 수 없습니다. cause :: %s", super.classType.getSimpleName(), e.getMessage()));
            throw new ExcelReadException();
        }
    }

    @Override
    protected void setTargetData(T data, String cellValue, @Nullable MethodInvoker keyValue) {
        ExcelReflectUtils.invokeTargetData(keyValue, data, cellValue);
    }

}
