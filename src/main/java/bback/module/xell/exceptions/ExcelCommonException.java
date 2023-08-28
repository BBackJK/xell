package bback.module.xell.exceptions;

public class ExcelCommonException extends RuntimeException {

    public ExcelCommonException(String msg) {
        super(msg);
    }

    public ExcelCommonException(Throwable e) {
        super(e);
    }
}
