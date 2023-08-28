package bback.module.xell.reader;

import bback.module.xell.exceptions.ExcelReadException;
import java.util.List;

public interface GridExcelReader<T> extends ExcelReader {

    List<T> read() throws ExcelReadException;
    List<T> read(int headerSkipCount) throws ExcelReadException;
}
