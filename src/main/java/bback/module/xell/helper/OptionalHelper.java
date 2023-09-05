package bback.module.xell.helper;

import java.util.Optional;
import java.util.function.Consumer;

public final class OptionalHelper<T> {

    private final Optional<T> optional;

    private OptionalHelper() {
        this.optional = Optional.empty();
    }

    private OptionalHelper(Optional<T> optional) {
        this.optional = optional;
    }

    public static <T> OptionalHelper<T> of(Optional<T> optional) {
        return new OptionalHelper<>(optional);
    }

    public void ifPresent(Consumer<? super T> consumer) {
        this.optional.ifPresent(consumer);
    }

    public void ifPresentOrElse(Consumer<? super T> consumer, Runnable runnable) {
        if (this.optional.isPresent()) {
            consumer.accept(this.optional.get());
        } else {
            runnable.run();
        }
    }
}
