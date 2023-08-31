package bback.module.xell.reader;

import bback.module.xell.annotations.ExcelTitle;
import bback.module.xell.enums.ExcelMime;
import bback.module.xell.exceptions.ExcelReadException;
import bback.module.xell.logger.Log;
import bback.module.xell.logger.LogFactory;
import bback.module.xell.util.ExcelListUtils;
import bback.module.xell.util.ExcelReflectUtils;
import bback.module.xell.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.lang.Nullable;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

public abstract class AbstractGridExcelReader<T, K> implements GridExcelReader<T> {

    private static final Log LOGGER = LogFactory.getLog(AbstractGridExcelReader.class);
    protected final Class<T> classType;
    protected InputStream is;

    protected AbstractGridExcelReader(Class<T> classType) {
        this.classType = classType;
    }

    protected AbstractGridExcelReader(Class<T> clazz, String filename) throws IOException, ExcelReadException {
        this(clazz, Files.newInputStream(Paths.get(filename)));
    }

    protected AbstractGridExcelReader(Class<T> clazz, InputStream in) throws ExcelReadException {
        this.classType = clazz;
        this.is = in;
    }

    protected abstract List<K> makeHeaderValueList(Sheet sheet, int headerSkipCount);
    protected abstract T createTarget();
    protected abstract void setTargetData(T data, String cellValue, @Nullable K keyValue);

    @Override
    public void setFile(InputStream is) {
        this.is = is;
    }

    @Override
    public List<T> read() throws ExcelReadException {
        return read(ExcelUtils.DEFAULT_HEADER_SKIP_COUNT);
    }

    @Override
    public List<T> read(int headerSkipCount) throws ExcelReadException {
        if (this.is == null) {
            throw new ExcelReadException("파일을 찾을 수 없습니다.");
        }
        try (Workbook wb = WorkbookFactory.create(this.is)) {
            Sheet sheet = wb.getSheetAt(0);
            int totalRowCount = sheet.getPhysicalNumberOfRows();
            List<T> resultList = new ArrayList<>(totalRowCount - headerSkipCount); // 총 row 수 - header row 수
            List<K> keyValues = this.makeHeaderValueList(sheet, headerSkipCount);
            resultList.addAll(this.makeDataList(sheet, headerSkipCount, keyValues));
            return resultList;
        } catch (Exception e) {
            throw new ExcelReadException(e);
        }
    }

    protected List<T> makeDataList(Sheet sheet, int headerSkipCount, List<K> keyValues) {
        List<T> resultList = new ArrayList<>();
        int totalRowCount = sheet.getPhysicalNumberOfRows();
        for (int i = headerSkipCount; i < totalRowCount; i++) {   // row 반복
            Row row = sheet.getRow(i);
            if (row == null) continue;

            int totalColumnCount = row.getPhysicalNumberOfCells();

            T data = this.createTarget();

            for (int j = 0; j < totalColumnCount; j++) {    // row 의 cell 반복
                Cell cell = row.getCell(j);
                if (cell == null) continue;
                String cellValue = this.getCellValue(cell);
                K keyValue = ExcelListUtils.getOrSupply(keyValues, j);
                this.setTargetData(data, cellValue, keyValue);
            }

            resultList.add(data);
        }
        return resultList;
    }

    public static void sampleDownload(HttpServletResponse response, String filename, Class<?> clazz) throws ExcelReadException {
        List<Field> fields = ExcelReflectUtils.filterFieldByExcelAnnotation(clazz, ExcelTitle.class);
        if ( fields.isEmpty() ) {
            throw new ExcelReadException(String.format("%s 클래스의 ExcelTitle 어노테이션이 설정 되어 있지 않습니다.", clazz.getSimpleName()));
        }
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            wb.setCompressTempFiles(true);

            SXSSFSheet sheet = wb.createSheet(ExcelUtils.DEFAULT_SHEET_NAME);

            // 헤더 생성
            Row header = sheet.createRow(0);
            int columnSize = fields.size();
            for (int i=0; i<columnSize;i++) {
                Field field = fields.get(i);
                ExcelTitle exceltitle = field.getAnnotation(ExcelTitle.class);
                Cell cell = header.createCell(i);
                cell.setCellValue(exceltitle.value());
            }

            // 바디 생성
            Row body = sheet.createRow(1);
            for (int i=0; i<columnSize;i++) {
                Field field = fields.get(i);
                ExcelTitle exceltitle = field.getAnnotation(ExcelTitle.class);
                Cell cell = body.createCell(i);
                String sampleValue = exceltitle.sample();
                cell.setCellValue(sampleValue);
            }

            // 파일 내리기.
            response.setContentType(ExcelUtils.isXlsx( filename ) ? ExcelMime.XLS.toString() : ExcelMime.XLSX.toString());
            response.setHeader(ExcelUtils.CONTENT_DISPOSITION, "attachment;filename=" + URLEncoder.encode(filename, "UTF-8") + ";");
            wb.write(response.getOutputStream());

        } catch (IOException e) {
            throw new ExcelReadException("샘플 다운로드하는데 실패하였습니다.");
        }
    }

    protected String getCellValue(Cell cell) {
        String result = null;
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getStringCellValue();
                break;
            case NUMERIC:
                if ( DateUtil.isCellDateFormatted(cell) ) { // date format
                    LocalDateTime ldt = cell.getLocalDateTimeCellValue();
                    result = ldt.toString();
                } else {
                    result = String.format("%.0f", cell.getNumericCellValue());
                }
                break;
            case BLANK:
                result = "";
                break;
            default:
                LOGGER.warn("허용되지 않은 엑셀 데이터 유형입니다. 파일을 확인해주세요. CellType :: " + cell.getCellType());
        }

        return result;
    }
}
