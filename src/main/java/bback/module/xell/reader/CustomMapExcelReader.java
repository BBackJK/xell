package bback.module.xell.reader;

import org.apache.poi.ss.usermodel.Cell;

public class CustomMapExcelReader extends MapExcelReader {

    @Override
    protected String getCellValue(Cell cell) {
        String result = null;
        switch (cell.getCellType()) {
            // case custom 구현
        }

        return result;
    }
}
