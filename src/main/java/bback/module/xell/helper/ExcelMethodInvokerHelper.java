package bback.module.xell.helper;

import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;
import org.springframework.util.MethodInvoker;

import java.lang.annotation.Annotation;
import java.util.Optional;

public class ExcelMethodInvokerHelper<T extends Annotation> implements Comparable<ExcelMethodInvokerHelper<T>> {

    @NonNull
    private final MethodInvoker invoker;

    @Nullable
    private final T annotation;
    private final int sort;

    public ExcelMethodInvokerHelper(MethodInvoker invoker, T annotation, int sort) {
        this.invoker = invoker;
        this.annotation = annotation;
        this.sort = sort;
    }

    public Optional<T> getAnnotation() {
        return Optional.ofNullable(this.annotation);
    }

    @Override
    public int compareTo(ExcelMethodInvokerHelper o) {
        if (o.sort < sort) {
            return 1;
        } else if ( o.sort > sort) {
            return -1;
        }
        return 0;
    }

    @Override
    public int hashCode() {
        return super.hashCode();
    }

    @Override
    public boolean equals(Object obj) {
        return super.equals(obj);
    }

    public MethodInvoker getInvoker() {
        return this.invoker;
    }
}
