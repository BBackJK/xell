package bback.module.xell.logger.empty;

import bback.module.xell.logger.Log;

public class NoLoggingImpl implements Log {

    public NoLoggingImpl(String clazz) {
        // ignore...
    }

    @Override
    public boolean isDebugEnabled() {
        return false;
    }

    @Override
    public boolean isTraceEnabled() {
        return false;
    }

    @Override
    public void error(String s, Throwable e) {
        // ignore...
    }

    @Override
    public void error(String s) {
        // ignore...
    }

    @Override
    public void debug(String s) {
        // ignore...
    }

    @Override
    public void trace(String s) {
        // ignore...
    }

    @Override
    public void warn(String s) {
        // ignore...
    }
}
