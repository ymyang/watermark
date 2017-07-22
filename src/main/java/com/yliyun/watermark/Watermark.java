package com.yliyun.watermark;

public class Watermark {
	
	public static void main(String[] args) {
		if (args == null || args.length < 3) {
			System.exit(1);
		}
		
		String src = args[0];
		String target = args[1];
		String text = args[2];
		
		try {
			boolean b = watermark(src, target, text);
			if (b) {
				System.exit(0);
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		
		System.exit(1);
	}

	public static boolean watermark(String src, String target, String text) throws Exception {
		String ext = getFileExt(src);
		if ("pdf".equalsIgnoreCase(ext)) {
			WatermarkPdf.watermark(src, target, text);
			return true;
		} else if ("docx".equalsIgnoreCase(ext)) {
			WatermarkOffice.word(src, target, text);
			return true;
		} else if ("xlsx".equalsIgnoreCase(ext)) {
			WatermarkOffice.excel2007(src, target, text);
			return true;
		} else if ("xls".equalsIgnoreCase(ext)) {
			WatermarkOffice.excel2003(src, target, text);
			return true;
		}
		return false;
	}
	
	private static String getFileExt(String file) {
		int index = file.lastIndexOf(".");
		if (index != -1) {
			return file.substring(index + 1);
		}
		return "";
	}
}
