package bback.module.xell.util;

import org.springframework.lang.Nullable;

import java.util.List;
import java.util.function.Supplier;

public final class ExcelListUtils {
    private ExcelListUtils() {
        throw new UnsupportedOperationException("This is a utility class and cannot be instantiated");
    }

    @Nullable
    public static <T> T getOrSupply(List<T> list, int idx) {
        return getOrSupply(list, idx, null);
    }

    @Nullable
    public static <T> T getOrSupply(List<T> list, int idx, Supplier<T> supplier) {
        try {
            return list.get(idx);
        } catch (IndexOutOfBoundsException ignore) {
            if (supplier == null) {
                return null;
            }
            T result = supplier.get();
            list.add(result);
            return result;
        } catch (Exception e) {
            throw e;
        }
    }

    public static boolean isEmpty(List<?> list) {
        return list == null || list.isEmpty();
    }
}
