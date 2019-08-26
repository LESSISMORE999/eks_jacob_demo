package com.eks.utils;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class FileUtils {
    public static String getProjectPath(){
        return System.getProperty("user.dir");
    }
    public static String generatePathBaseProjectPath(String relativePathString){
        return getProjectPath() + File.separator + relativePathString;
    }
    public static List<String> convertFileToList(String filePath) throws Exception {
        return convertFileToList(filePath, "UTF-8");
    }
    public static List<String> convertFileToList(String filePath, String encoding) throws Exception {
        List<String> stringList = new ArrayList<String>();
        File file = new File(filePath);
        if (file.isFile() && file.exists()) {
            InputStreamReader inputStreamReader = new InputStreamReader(new FileInputStream(file), encoding);
            BufferedReader bufferedReader = new BufferedReader(inputStreamReader);
            String lineTxt = null;
            while ((lineTxt = bufferedReader.readLine()) != null) {
                stringList.add(lineTxt);
            }
            bufferedReader.close();
            inputStreamReader.close();
        } else {
            throw new Exception("The specified file could not be found");
        }
        return stringList;
    }
    public static String getContentStringBaseProjectPath(String relativePathString) throws Exception {
        return convertFileToContentString(generatePathBaseProjectPath(relativePathString));
    }
    public static String convertFileToContentString(String filePathString) throws Exception {
        List<String> everyLineListFromFile = convertFileToList(filePathString);
        StringBuilder stringBuilder = new StringBuilder();
        for(String string : everyLineListFromFile){
            stringBuilder.append(string);
        }
        return stringBuilder.toString();
    }
    public static void writeStringToFile(String filePathString, String writeString, Boolean appendBoolean) throws Exception {
        if (filePathString == null || "".equals(filePathString)){
            throw new RuntimeException("FilePathString should not be null or empty.");
        }
        if (writeString == null || "".equals(writeString)){
            throw new RuntimeException("WriteString should not be null or empty.");
        }
        FileOutputStream fileOutputStream = null;
        try {
            File file = getFileByFilePath(filePathString);
            fileOutputStream = new FileOutputStream(file,appendBoolean);
            fileOutputStream.write(writeString.getBytes());
        } finally {
            if (fileOutputStream != null){
                fileOutputStream.close();
            }
        }
    }
    public static File getFileByFilePath(String filePathString) throws IOException {
        File file = new File(filePathString);
        if (!file.exists()) {
            File dirFile = new File(file.getParent());
            if (!dirFile.exists()) {
                boolean mkdirsBoolean = dirFile.mkdirs();
                if (!mkdirsBoolean) {
                    throw new RuntimeException("Fail to create dir.");
                }
            }
            boolean newFileBoolean = file.createNewFile();
            if (!newFileBoolean) {
                throw new RuntimeException("Fail to create new file.");
            }
        }
        return file;
    }
}
