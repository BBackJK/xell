package bback.module.xell.writer;

import bback.module.xell.exceptions.ExcelWriteException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Collections;
import java.util.List;

public class ListGridExcelWriter extends AbstractGridExcelWriter<List<Object>> {

    private final List<List<String>> header;

    public ListGridExcelWriter(List<List<String>> header) {
        this(header, Collections.emptyList());
    }

    public ListGridExcelWriter(List<List<String>> header, List<List<Object>> dataList) {
        super((Class<List<Object>>) dataList.getClass(), dataList);
        this.header = header;
    }

    @Override
    protected void writeHeader(SXSSFWorkbook wb, SXSSFSheet sheet) throws ExcelWriteException {
        List<List<String>> headerList = this.header == null ? Collections.emptyList() : this.header;
        int headerRowCount = headerList.size();
        for (int i=0; i<headerRowCount; i++) {
            Row row = sheet.createRow(i);
            List<String> headerCellList = headerList.get(i);
            int headerCellCount = headerCellList.size();
            for (int j=0; j<headerCellCount;j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(headerCellList.get(i));
            }
        }
    }

    @Override
    protected void writeBody(SXSSFWorkbook wb, SXSSFSheet sheet) throws ExcelWriteException {
        List<List<Object>> bodyList = this.dataList == null ? Collections.emptyList() : this.dataList;
        int lastHeaderRowIndex = sheet.getLastRowNum();
        int bodyRowCount = bodyList.size();
        int rowNum = lastHeaderRowIndex + 1;
        for (int i=0; i<bodyRowCount; i++) {
            Row row = sheet.createRow(rowNum++);
            List<Object> bodyCellList = bodyList.get(i);
            int bodyCellCount = bodyCellList.size();
            for (int j=0; j<bodyCellCount;j++) {
                Cell cell = row.createCell(j);
                Object cellValue = bodyCellList.get(i);
                cell.setCellValue(cellValue == null ? "" : String.valueOf(cellValue));
            }
        }
    }
}
