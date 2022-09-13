package com.github.stl4j.awesomepdf.base.util;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Various utility functions for easier loading and finding the system fonts.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @since 0.0.1
 */
public class FontUtils {

    /**
     * The OS fonts directory mapping, support macOS, Linux, and Windows platforms.
     * The key of this map is the OS platform, and the value is the OS fonts directory.
     */
    private static final Map<String, String> FONTS_DIRECTORY_MAP = new HashMap<>(3);

    /**
     * The default font mapping, we will use the default font finally if no valid OS font matched by the given font name.
     * The key of this map is the OS platform, and the value is the file path of default font.
     */
    private static final Map<String, String> DEFAULT_FONT_MAP = new HashMap<>(3);

    /**
     * Current OS platform
     */
    private static final String CURRENT_OS_PLATFORM = System.getProperty("os.name");

    /**
     * The font formats that itext pdf library supported now.
     */
    private static final String[] SUPPORTED_FONT_FORMATS = {".ttf", ".otf", ".ttc"};

    /**
     * All the font files in current OS platform.
     */
    private static final List<File> SYSTEM_FONT_FILES;

    // Initialize some variables and load OS font files only once.
    static {
        FONTS_DIRECTORY_MAP.put("Mac OS X", "/System/Library/Fonts");
        DEFAULT_FONT_MAP.put("Mac OS X", "/System/Library/Fonts/Supplemental/Songti.ttc");
        final String systemFontDirectory = FONTS_DIRECTORY_MAP.get(CURRENT_OS_PLATFORM);
        SYSTEM_FONT_FILES = FileUtils.listFilesRecursively(systemFontDirectory);
    }

    private FontUtils() {
        // An utility class should not be instantiated by external classes.
    }

    /**
     * Find the system font file path by the given font name.
     *
     * @param fontFullName The given font name.
     * @return The system font file path.
     */
    public static String findSystemFontPath(String fontFullName) {
        final String systemFontDirectory = FONTS_DIRECTORY_MAP.get(CURRENT_OS_PLATFORM);
        String result = null;

        for (String fontFormat : SUPPORTED_FONT_FORMATS) {
            final String fontFileName = fontFullName + fontFormat;
            Path fontFilePath = Paths.get(systemFontDirectory, fontFileName);
            result = Files.exists(fontFilePath) ? fontFilePath.toString() : tryGuessFontFilePath(fontFullName);
            if (result != null) {
                break;
            }
        }

        if (result == null) {
            return DEFAULT_FONT_MAP.get(CURRENT_OS_PLATFORM);
        }

        // Process .ttc font format, get the 0th font by default.
        if (result.endsWith(SUPPORTED_FONT_FORMATS[SUPPORTED_FONT_FORMATS.length - 1])) {
            return result + Characters.COMMA + '0';
        }

        return result;
    }

    /**
     * We will try to guess the font file path if can't find it by the given font name.
     * Here's how to do it: generally the length of font full name will be greater than the
     * length of font file name, so we can assume that the font full name will contain
     * the font file name. Of course this is not guaranteed, so we still return {@code null}
     * if there is no valid font file found.
     *
     * @param fontFullName The full name of the font.
     * @return The absolute path of the font file.
     */
    private static String tryGuessFontFilePath(String fontFullName) {
        for (File fontFile : SYSTEM_FONT_FILES) {
            final String fontFileBaseName = FileUtils.getFileBaseName(fontFile);
            // Generally the fontFullName.length() should be greater than fontFileBaseName.length(), that's why we match it like this.
            if (fontFullName.toLowerCase().contains(fontFileBaseName.toLowerCase())) {
                return fontFile.getAbsolutePath();
            }
        }
        return null;
    }

}
