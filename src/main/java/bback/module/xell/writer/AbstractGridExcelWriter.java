package bback.module.xell.writer;

import bback.module.xell.enums.ExcelMime;
import bback.module.xell.exceptions.ExcelWriteException;
import bback.module.xell.logger.Log;
import bback.module.xell.logger.LogFactory;
import bback.module.xell.util.ExcelUtils;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.List;

public abstract class AbstractGridExcelWriter<T> implements GridExcelWriter<T> {

    private static final Log LOGGER = LogFactory.getLog(AbstractGridExcelWriter.class);
    protected final Class<T> classType;
    protected List<T> dataList;

    protected AbstractGridExcelWriter(Class<T> classType) {
        this(classType, Collections.emptyList());
    }

    protected AbstractGridExcelWriter(Class<T> classType, List<T> dataList) {
        this.classType = classType;
        this.dataList = dataList;
    }

    protected abstract void writeHeader(SXSSFWorkbook wb, SXSSFSheet sheet) throws ExcelWriteException;
    protected abstract void writeBody(SXSSFWorkbook wb, SXSSFSheet sheet) throws ExcelWriteException;


    @Override
    public void setDataList(List<T> dataList) {
        this.dataList = dataList;
    }

    @Override
    public void generate(String filePath) throws ExcelWriteException {
        generate(filePath, ExcelUtils.DEFAULT_SHEET_NAME);
    }

    @Override
    public void generate(String filePath, String sheetName) throws ExcelWriteException {
        try {
            ExcelUtils.makeExcelFilePath(filePath);
            this.write(Files.newOutputStream(Paths.get(filePath)), sheetName);
        } catch (IOException e) {
            throw new ExcelWriteException(e);
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
            this.write(response.getOutputStream(), sheetName);
        } catch (IOException e) {
            throw new ExcelWriteException(e);
        }
    }

    private void write(OutputStream out, String sheetName) {
        try (SXSSFWorkbook wb = new SXSSFWorkbook(ExcelUtils.WINDOW_SIZE)) {
            wb.setCompressTempFiles(true);

            SXSSFSheet sheet = wb.createSheet(sheetName);
            sheet.setRandomAccessWindowSize(ExcelUtils.WINDOW_SIZE);

            this.writeHeader(wb, sheet); // 헤더 생성
            this.writeBody(wb, sheet); // 바디 생성

            wb.write(out);
            out.close();
            wb.dispose();
        } catch (ExcelWriteException | IOException e) {
            LOGGER.error("엑셀을 만드는데 실패하였습니다. cause :: " + e.getMessage(), e);
            throw new ExcelWriteException();
        }
    }
}
