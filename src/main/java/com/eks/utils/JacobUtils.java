package com.eks.utils;

public class JacobUtils {
    public static void addJavaLibraryPath(){
        try {
            SystemUtils.addJavaLibraryPath(FileUtils.generatePathBaseProjectPath("extra/jacob"));
        } catch (Exception e) {
            throw new RuntimeException("Fail to add java.library.path",e);
        }
    }
}
