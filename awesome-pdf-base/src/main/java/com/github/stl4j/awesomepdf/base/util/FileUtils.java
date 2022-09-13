package com.github.stl4j.awesomepdf.base.util;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Various utility functions for easier processing files.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @since 0.0.1
 */
public class FileUtils {

    private FileUtils() {
        // An utility class should not be instantiated by external classes.
    }

    /**
     * List and return all the files in the given directory, including nested directories.
     *
     * @param directory The directory to search for files.
     * @return The list of files in the given directory.
     */
    public static List<File> listFilesRecursively(String directory) {
        File rootDirectory = new File(directory);
        File[] files = rootDirectory.listFiles();

        if (files == null || files.length == 0) {
            return Collections.emptyList();
        }

        List<File> result = new ArrayList<>();
        for (File file : files) {
            if (file.isDirectory()) {
                List<File> subFiles = listFilesRecursively(file.getAbsolutePath());
                result.addAll(subFiles);
            } else if (file.isFile()) {
                result.add(file);
            }
        }
        return result;
    }

    /**
     * Get the base name of the given file, in other words, remove the extension from its name.
     *
     * @param file The file to get the base name.
     * @return The base name of the given file.
     */
    public static String getFileBaseName(File file) {
        final String fileName = file.getName();
        final int lastDotIndex = fileName.lastIndexOf(Characters.DOT);
        return lastDotIndex == -1 ? fileName : fileName.substring(0, lastDotIndex);
    }

}
