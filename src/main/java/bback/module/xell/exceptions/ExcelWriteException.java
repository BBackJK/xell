package bback.module.xell.exceptions;

public class ExcelWriteException extends ExcelCommonException {

    public ExcelWriteException() {
        super("엑셀 파일을 만드는데 실패하였습니다.");
    }

    public ExcelWriteException(Throwable e) {
        super(e);
    }

    public ExcelWriteException(String m) {
        super(m);
    }
}
