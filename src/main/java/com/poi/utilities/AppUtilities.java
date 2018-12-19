package com.poi.utilities;


import java.io.InputStream;

public class AppUtilities {
	public final static String CLASS_KEY = "AppUtilities";

	/**
	 * Used For get InputStream
	 * 
	 * @param fileName
	 * @return
	 */
	public static InputStream getFileName(String fileName) {
		final String METHOD_KEY = "getFileName";
		ClassLoader classloader = Thread.currentThread().getContextClassLoader();
		InputStream is = classloader.getResourceAsStream(fileName);
		return is;
	}
	// generate excel workbook - extension based on defined output file
	public static String getExtension(String fileName) {
		String extension = null;
		if (fileName == null) {
			return extension;
		} else {
			int i = fileName.lastIndexOf('.');
			if (i > 0) {
				extension = fileName.substring(i + 1);
			}
			// uing Apache Commons IO
			// extension = FilenameUtils.getExtension(fileName);
		}
		return extension;
	}
}
