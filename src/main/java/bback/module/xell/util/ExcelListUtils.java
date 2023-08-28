package bback.module.xell.util;

import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;

import java.util.List;
import java.util.function.Supplier;

@UtilityClass
@Slf4j
public class ExcelListUtils {

    public <T> T getOrSupply(List<T> list, int idx) {
        return getOrSupply(list, idx, null);
    }

    public <T> T getOrSupply(List<T> list, int idx, Supplier<T> supplier) {
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
}
