package bback.module.xell.reader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.lang.Nullable;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MapExcelReader extends AbstractGridExcelReader<Map, String> {

    private final Map<String, String> keyMap;

    public MapExcelReader() {
        this((Map<String, String>) null);
    }

    public MapExcelReader(Map<String, String> keyMap) {
        super(Map.class);
        this.keyMap = keyMap;
    }

    public MapExcelReader(String filePath) throws IOException {
        this(Files.newInputStream(Paths.get(filePath)));
    }

    public MapExcelReader(InputStream is) {
        super(Map.class, is);
        this.keyMap = null;
    }

    @Override
    protected List<String> makeHeaderValueList(Sheet sheet, int headerSkipCount) {
        List<String> headers = new ArrayList<>();
        for (int i=0; i<headerSkipCount;i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            int totalColumns = row.getPhysicalNumberOfCells();
            for (int j = 0; j < totalColumns; j++) {    // row 의 cell 반복
                Cell cell = row.getCell(j);
                if (cell == null) continue;
                String cellValue = cell.getStringCellValue();
                String keyMapValue = null;
                if (this.keyMap != null) {
                    keyMapValue = this.keyMap.get(cellValue);
                }
                headers.add(keyMapValue == null ? cellValue : keyMapValue);
            }
        }
        return headers;
    }

    @Override
    protected Map createTarget() {
        return new HashMap<String, String>();
    }

    @Override
    protected void setTargetData(Map data, String cellValue, @Nullable String keyValue) {
        if ( data != null && keyValue != null) {
            data.put(keyValue, cellValue);
        }
    }
}
