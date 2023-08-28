package bback.module.xell.exceptions;

public class ExcelReadException extends ExcelCommonException {

    public ExcelReadException() {
        super("엑셀 정보를 읽을 수 없습니다.");
    }

    public ExcelReadException(Throwable e) {
        super(e);
    }
    public ExcelReadException(String m) {
        super(m);
    }
}
