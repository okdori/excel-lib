package com.okdori.utils;

import java.lang.reflect.Field;

/**
 * packageName    : com.okdori.utils
 * fileName       : TypeUtils
 * author         : okdori
 * date           : 2024. 12. 20.
 * description    :
 */

public class TypeUtils {
    public static boolean isPrimitiveOrSimpleType(Field field) {
        Class<?> type = field.getType();
        return type.isPrimitive()
                || type.equals(String.class)
                || java.time.temporal.Temporal.class.isAssignableFrom(type)
                || Number.class.isAssignableFrom(type);
    }
}
