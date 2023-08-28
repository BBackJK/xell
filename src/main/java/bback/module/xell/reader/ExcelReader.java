package bback.module.xell.reader;

import bback.module.xell.exceptions.ExcelReadException;
import java.io.InputStream;

public interface ExcelReader {

    void setFile(InputStream is) throws ExcelReadException;
}
