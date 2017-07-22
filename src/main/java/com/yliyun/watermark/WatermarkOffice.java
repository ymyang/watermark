package com.yliyun.watermark;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFTextbox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public final class WatermarkOffice {

	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {
		try {
			String src = "D:\\watermark\\test.xls";
			String target = "D:\\watermark\\test-watermark.xls";
			String text = "YLIYUN";
			excel2003(src, target, text);
			System.out.println("ok");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public static void word(String src, String target, String text) throws IOException {
		XWPFDocument doc = null;
		OutputStream out = null;
		try {
			doc = new XWPFDocument(new FileInputStream(src));
			XWPFHeaderFooterPolicy footer = new XWPFHeaderFooterPolicy(doc);
			footer.createWatermark(text);
			out = new FileOutputStream(target);
			doc.write(out);
		} finally {
			if (doc != null) {
				doc.close();
			}
			if (out != null) {
				out.close();
			}
		}

	}

	public static void excel2003(String src, String target, String text)
			throws IOException, EncryptedDocumentException, InvalidFormatException {
		HSSFWorkbook wb = null;
		OutputStream out = null;
		try {
			InputStream input = new FileInputStream(src);

			wb = (HSSFWorkbook) WorkbookFactory.create(input);
			HSSFSheet sheet = null;

			int sheetNumbers = wb.getNumberOfSheets();

			// sheet
			for (int i = 0; i < sheetNumbers; i++) {
				sheet = wb.getSheetAt(i);
				// sheet.createDrawingPatriarch();

				HSSFPatriarch dp = sheet.createDrawingPatriarch();
				HSSFClientAnchor anchor = new HSSFClientAnchor(0, 255, 550, 0, (short) 0, 1, (short) 6, 5);

				// HSSFComment comment = dp.createComment(anchor);
				HSSFTextbox txtbox = dp.createTextbox(anchor);

				HSSFRichTextString rtxt = new HSSFRichTextString(text);
				HSSFFont draftFont = (HSSFFont) wb.createFont();
				// 水印颜色
				draftFont.setColor((short) 55);
				draftFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				// 字体大小
				draftFont.setFontHeightInPoints((short) 30);
				draftFont.setFontName("Verdana");
				rtxt.applyFont(draftFont);
				txtbox.setString(rtxt);
				// 倾斜度
				txtbox.setRotationDegree((short) 315);
				txtbox.setLineWidth(600);
				txtbox.setLineStyle(HSSFShape.LINESTYLE_NONE);
				txtbox.setNoFill(true);
			}
			
			out = new FileOutputStream(target);
			wb.write(out);
		} finally {
			if (wb != null) {
				wb.close();
			}
			if (out != null) {
				out.close();
			}
		}

	}

	public static void excel2007(String src, String target, String text)
			throws IOException, EncryptedDocumentException, InvalidFormatException {
		XSSFWorkbook wb = null;
		OutputStream out = null;
		try {
			InputStream input = new FileInputStream(src);
			wb = (XSSFWorkbook) WorkbookFactory.create(input);
			
			XSSFSheet sheet = null;
			int sheetNumbers = wb.getNumberOfSheets();
			for (int i = 0; i < sheetNumbers; i++) {
				sheet = wb.getSheetAt(i);
				XSSFDrawing dp = sheet.createDrawingPatriarch();
				XSSFClientAnchor anchor = new XSSFClientAnchor(0, 550, 550, 0, (short) 0, 1, (short) 6, 5);
				XSSFTextBox txtbox = dp.createTextbox(anchor);
				XSSFRichTextString rtxt = new XSSFRichTextString(text);
				XSSFFont draftFont = (XSSFFont) wb.createFont();
				draftFont.setColor((short) 55);
				draftFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
				draftFont.setFontHeightInPoints((short) 30);
				draftFont.setFontName("Verdana");
				rtxt.applyFont(draftFont);
				txtbox.setText(rtxt);
				// 倾斜度
				txtbox.setLineWidth(600);
				txtbox.setLineStyle(HSSFShape.LINESTYLE_NONE);
				txtbox.setNoFill(true);
			}
			
			out = new FileOutputStream(target);
			wb.write(out);
		} finally {
			if (wb != null) {
				wb.close();
			}
			if (out != null) {
				out.close();
			}
		}
	}

}