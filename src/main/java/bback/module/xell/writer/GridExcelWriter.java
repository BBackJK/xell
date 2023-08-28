package bback.module.xell.writer;

import java.util.List;

public interface GridExcelWriter<T> extends ExcelWriter {
    void setDataList(List<T> dataList);
}
