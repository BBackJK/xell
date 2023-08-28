package bback.module.xell.writer;

import bback.module.xell.exceptions.ExcelWriteException;

import javax.servlet.http.HttpServletResponse;

public interface ExcelWriter {

    void generate(String filePath) throws ExcelWriteException;
    void generate(String filePath, String sheetName) throws ExcelWriteException;
    void download(HttpServletResponse response, String filename) throws ExcelWriteException;
    void download(HttpServletResponse response, String filename, String sheetName) throws ExcelWriteException;
}
