package bback.module.xell.writer;

import bback.module.xell.annotations.ExcelPointer;
import bback.module.xell.enums.ExcelMime;
import bback.module.xell.exceptions.ExcelWriteException;
import bback.module.xell.helper.ExcelMethodInvokerHelper;
import bback.module.xell.helper.ExcelSheetInfo;
import bback.module.xell.logger.Log;
import bback.module.xell.logger.LogFactory;
import bback.module.xell.util.ExcelReflectUtils;
import bback.module.xell.util.ExcelStringUtils;
import bback.module.xell.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressBase;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.lang.Nullable;
import org.springframework.util.MethodInvoker;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public abstract class AbstractPointExcelWriter<T> implements PointExcelWriter<T> {
    private static final Log LOGGER = LogFactory.getLog(AbstractPointExcelWriter.class);
    protected final Class<T> classType;
    protected final Map<String, ExcelMethodInvokerHelper<ExcelPointer>> excelMethodInvokerHelperMap;
    protected InputStream in;

    @Nullable
    protected T data;

    protected String startPattern = "\\$\\{\\{";
    protected String endPattern = "}}";

    protected AbstractPointExcelWriter(Class<T> classType) {
        this(classType, null);
    }

    protected AbstractPointExcelWriter(Class<T> classType, T data) {
        this.classType = classType;
        this.data = data;
        this.excelMethodInvokerHelperMap = this.getExcelMethodInvokerHelperMap(classType);
    }

    @Override
    public void setData(T data) {
        this.data = data;
    }

    @Override
    public void setInputStream(InputStream in) {
        this.in = in;
    }

    @Override
    public void setStartPattern(String pattern) {
        this.startPattern = pattern;
    }

    @Override
    public void setEndPattern(String pattern) {
        this.endPattern = pattern;
    }

    @Override
    public void generate(String filePath) throws ExcelWriteException {
        generate(filePath, ExcelUtils.DEFAULT_SHEET_NAME);
    }

    @Override
    public void generate(String filePath, String sheetName) throws ExcelWriteException {
        try {
            ExcelUtils.makeExcelFilePath(filePath);
            this.write(Files.newOutputStream(Paths.get(filePath)));
        } catch (IOException e) {
            throw new ExcelWriteException();
        }
    }

    @Override
    public void download(HttpServletResponse response, String filename) throws ExcelWriteException {
        download(response, filename, ExcelUtils.DEFAULT_SHEET_NAME);
    }

    @Override
    public void download(HttpServletResponse response, String filename, String sheetName) throws ExcelWriteException {
        try {
            // 파일 내리기.
            response.setContentType(ExcelUtils.isXlsx( filename ) ? ExcelMime.XLSX.toString() : ExcelMime.XLS.toString());
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(filename, "UTF-8") + ";");
            this.write(response.getOutputStream());
        } catch (IOException e) {
            throw new ExcelWriteException(e);
        }
    }

    protected abstract void setRow(Sheet sourceSheet, Sheet targetSheet, Row source, Row target);
    protected abstract void setCell(Sheet sourceSheet, Sheet targetSheet, Cell source, Cell target);

    private Map<String, ExcelMethodInvokerHelper<ExcelPointer>> getExcelMethodInvokerHelperMap(Class<?> classType) {
        Map<String, ExcelMethodInvokerHelper<ExcelPointer>> result = new HashMap<>();
        List<Field> excelPointFieldList = ExcelReflectUtils.filterFieldByExcelAnnotation(classType, ExcelPointer.class);
        List<Method> excelPointMethodList = ExcelReflectUtils.filterMethodByExcelAnnotation(classType, ExcelPointer.class);

        int fieldCount = excelPointFieldList.size();
        for (int i=0;i<fieldCount;i++) {
            Field f = excelPointFieldList.get(i);
            ExcelPointer excelPointer = f.getAnnotation(ExcelPointer.class);
            MethodInvoker invoker = new MethodInvoker();
            invoker.setTargetMethod(ExcelStringUtils.toGetterByField(f));
            result.put(
                    excelPointer.value()
                    , new ExcelMethodInvokerHelper<>(invoker, excelPointer, 0)
            );
        }

        int methodCount = excelPointMethodList.size();
        for (int i=0;i<methodCount;i++) {
            Method m = excelPointMethodList.get(i);
            ExcelPointer excelPointer = m.getAnnotation(ExcelPointer.class);
            MethodInvoker invoker = new MethodInvoker();
            invoker.setTargetMethod(m.getName());
            result.put(
                    excelPointer.value()
                    , new ExcelMethodInvokerHelper<>(invoker, excelPointer, 0)
            );
        }

        return result;
    }

    private void write(OutputStream out) {
        if (this.in == null) {
            throw new ExcelWriteException("엑셀 양식 파일에 대한 정보가 없습니다.");
        }
        List<ExcelSheetInfo> excelSheetInfoList = this.getSourceExcelSheetInfo(this.in);

        try (SXSSFWorkbook targetWorkbook = new SXSSFWorkbook(ExcelUtils.WINDOW_SIZE)) {
            int sourceSheetCount = excelSheetInfoList.size();
            if (sourceSheetCount < 1) {
                throw new ExcelWriteException("원본 엑셀 파일이 유효하지 않습니다.");
            }
            for (int i=0; i<sourceSheetCount;i++) {
                ExcelSheetInfo excelSheetInfo = excelSheetInfoList.get(i);
                Sheet sourceSheet = excelSheetInfo.getSheet();
                List<Row> sourceRowList = excelSheetInfo.getRowList();
                int maxRowIndex = excelSheetInfo.getMaxRowIndex();
                int maxColumnIndex = excelSheetInfo.getMaxColumnIndex();

                SXSSFSheet targetSheet = targetWorkbook.createSheet(sourceSheet.getSheetName());
                targetSheet.setRandomAccessWindowSize(ExcelUtils.WINDOW_SIZE);

                for (int j=0; j<=maxRowIndex;j++) {
                    Row sourceRow = sourceRowList.get(j);
                    Row targetRow = targetSheet.createRow(j);

                    // row set
                    this.setRow(sourceSheet, targetSheet, sourceRow, targetRow);

                    for (int z=0; z<=maxColumnIndex;z++) {
                        Cell sourceCell = sourceRow.getCell(z); // @Nullable
                        Cell targetCell = targetRow.createCell(z);

                        // cell set
                        this.setCell(sourceSheet, targetSheet, sourceCell, targetCell);
                    }
                }

                List<CellRangeAddress> mergeList = excelSheetInfo.getMergeInfoList();
                if (!mergeList.isEmpty()) {
                    int mergeCellCount = mergeList.size();
                    for (int x=0; x< mergeCellCount;x++) {
                        targetSheet.addMergedRegion(mergeList.get(x));
                    }
                }
            }

            targetWorkbook.write(out);
            out.close();
            targetWorkbook.dispose();
        } catch (IOException e) {
            LOGGER.error(e.getMessage());
            throw new ExcelWriteException();
        }
    }

    private List<ExcelSheetInfo> getSourceExcelSheetInfo(InputStream in) {
        try (Workbook origin = WorkbookFactory.create(in)) {
            List<ExcelSheetInfo> excelSheetInfoList = new ArrayList<>();
            int sheetCount = origin.getNumberOfSheets();
            for (int i=0; i<sheetCount; i++) {
                Sheet sheet = origin.getSheetAt(i);

                List<CellRangeAddress> mergeInfoList = new ArrayList<>();
                List<Row> rowList = new ArrayList<>();
                int maxRowIndex = 0;
                int maxColumnIndex = 0;

                int mergedRegionCount = sheet.getNumMergedRegions();
                if ( mergedRegionCount > 0 ) {
                    mergeInfoList.addAll(sheet.getMergedRegions());
                    maxRowIndex = mergeInfoList.stream().mapToInt(CellRangeAddressBase::getLastRow).max().orElseGet(() -> 0);
                    maxColumnIndex = mergeInfoList.stream().mapToInt(CellRangeAddressBase::getLastColumn).max().orElseGet(() -> 0);
                }

                int physicalRowIndex = sheet.getPhysicalNumberOfRows() - 1;
                if (maxRowIndex < physicalRowIndex) {
                    i = 0; // 초기화 --> sheet 다시반복
                    maxRowIndex = physicalRowIndex;
                }
                for (int j=0; j<=maxRowIndex;j++) {
                    Row row = sheet.getRow(j);

                    int physicalColumnIndex = row.getPhysicalNumberOfCells() - 1;
                    if (maxColumnIndex < physicalColumnIndex) {
                        j=0;    // 초기화 --> row 다시반복
                        rowList.clear();
                        maxColumnIndex = physicalColumnIndex;
                    }
                    rowList.add(row);
                }

                excelSheetInfoList.add(
                        new ExcelSheetInfo(
                                sheet
                                , mergeInfoList
                                , rowList
                                , maxRowIndex
                                , maxColumnIndex
                        )
                );
            }
            return excelSheetInfoList;
        } catch (IOException e) {
            throw new ExcelWriteException();
        }
    }
}
