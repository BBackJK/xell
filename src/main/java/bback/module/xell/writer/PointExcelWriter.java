package bback.module.xell.writer;

import java.io.InputStream;

public interface PointExcelWriter<T> extends ExcelWriter {
    void setData(T data);
    void setInputStream(InputStream in);
    void setStartPattern(String pattern);
    void setEndPattern(String pattern);
}
