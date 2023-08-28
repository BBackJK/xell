package bback.module.xell.large;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.io.Writer;
import java.util.Calendar;

public class SpreadSheetWriter {
    public static final String XML_ENCODING = "UTF-8";
    private final Writer writer;
    private int rowNum;

    public SpreadSheetWriter(Writer out){
        writer = out;
    }

    public void beginSheet() throws IOException {
        writer.write("<?xml version=\"1.0\" encoding=\""+XML_ENCODING+"\"?>" +
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" );
        writer.write("<sheetData>\n");
    }

    public void endSheet() throws IOException {
        writer.write("</sheetData>");
        writer.write("</worksheet>");
    }

    /**
     * Insert a new row
     *
     * @param rownum 0-based row number
     */
    public void insertRow(int rownum) throws IOException {
        writer.write("<row r=\""+(rownum+1)+"\">\n");
        this.rowNum = rownum;
    }

    /**
     * Insert row end marker
     */
    public void endRow() throws IOException {
        writer.write("</row>\n");
    }

    public void createCell(int columnIndex, String value, int styleIndex) throws IOException {
        String ref = new CellReference(rowNum, columnIndex).formatAsString();
        writer.write("<c r=\""+ref+"\" t=\"inlineStr\"");
        if(styleIndex != -1) writer.write(" s=\""+styleIndex+"\"");
        writer.write(">");
        writer.write("<is><t>"+value+"</t></is>");
        writer.write("</c>");
    }

    public void createCell(int columnIndex, String value) throws IOException {
        createCell(columnIndex, value, -1);
    }

    public void createCell(int columnIndex, double value, int styleIndex) throws IOException {
        String ref = new CellReference(rowNum, columnIndex).formatAsString();
        writer.write("<c r=\""+ref+"\" t=\"n\"");
        if(styleIndex != -1) writer.write(" s=\""+styleIndex+"\"");
        writer.write(">");
        writer.write("<v>"+value+"</v>");
        writer.write("</c>");
    }

    public void createCell(int columnIndex, double value) throws IOException {
        createCell(columnIndex, value, -1);
    }

    public void createCell(int columnIndex, Calendar value, int styleIndex) throws IOException {
        createCell(columnIndex, DateUtil.getExcelDate(value, false), styleIndex);
    }
}
