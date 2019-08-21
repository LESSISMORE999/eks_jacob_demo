package com.eks.utils;

import java.lang.reflect.Field;

public class SystemUtils {
    public static void addJavaLibraryPath(String javaLibraryPathString) throws NoSuchFieldException, IllegalAccessException {
        String originalJavaLibraryPathString = System.getProperty("java.library.path");
        originalJavaLibraryPathString = javaLibraryPathString + ";" + originalJavaLibraryPathString;
        System.setProperty("java.library.path", originalJavaLibraryPathString);
        Field field = ClassLoader.class.getDeclaredField("sys_paths");
        field.setAccessible(true);
        field.set(null, null);
    }
}
