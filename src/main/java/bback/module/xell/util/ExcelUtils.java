package bback.module.xell.util;

import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;

@UtilityClass
@Slf4j
public class ExcelUtils {

    public final String CONTENT_DISPOSITION = "Content-Disposition";
    public final int WINDOW_SIZE = 100;
    public final int DEFAULT_HEADER_SKIP_COUNT = 1;
    public final String DEFAULT_SHEET_NAME = "Sheet";

    public final String EXT_XLS = "xls";
    public final String EXT_XLSX = "xlsx";

    public boolean isExcel(String filename) {
        String ext = getExtension(filename);
        return EXT_XLS.equals(ext) || EXT_XLSX.equals(ext);
    }

    public boolean isXlsx(String filename) {
        return filename != null && EXT_XLSX.equals(getExtension(filename));
    }

    public String getExtension(String filename) {
        return filename == null || filename.isEmpty() ? "" : filename.substring(filename.lastIndexOf(".") + 1);
    }

    public void makeExcelFilePath(String fileFullPath) throws IOException {
        if (fileFullPath == null || fileFullPath.isEmpty()) return;
        File file = new File(fileFullPath);
        if ( !file.exists() ) {
            int separateIndex = fileFullPath.lastIndexOf(File.separator);
            String path = fileFullPath.substring(0, separateIndex + 1);
            String filename = fileFullPath.substring(separateIndex + 1);

            if (!ExcelUtils.isExcel(filename)) {
                throw new IllegalArgumentException("파일 경로에 파일명이 포함되어있지 않습니다.");
            }

            File filePath = new File(path);
            boolean dirExist = filePath.exists() || filePath.mkdirs();
            if (dirExist) {
                boolean result = file.createNewFile();
                if (!result) {
                    throw new IOException("파일을 생성하는데에 실패하였습니다.");
                }
            }
        }
    }
}
