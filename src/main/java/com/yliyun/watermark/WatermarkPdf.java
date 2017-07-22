package com.yliyun.watermark;

import java.io.FileOutputStream;
import java.io.IOException;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

public class WatermarkPdf {

	public static void main(String[] args) {
		try {
			String src = "D:\\watermark\\test.pdf";
			String target = "D:\\watermark\\test-watermark.pdf";
			String text = "YLIYUN";
			watermark(src, target, text);
			System.out.println("ok");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
	
	public static void watermark(String src, String target, String text) throws IOException, DocumentException {
		PdfReader reader = null;
		PdfStamper pdfStamper = null;
		try {
			reader = new PdfReader(src);
			pdfStamper = new PdfStamper(reader, new FileOutputStream(target));
			
			addWatermark(pdfStamper, text);
		} finally {
			if (pdfStamper != null) {
				pdfStamper.close();
			}
		}
	}

	private static void addWatermark(PdfStamper pdfStamper, String watermark) throws DocumentException, IOException {
		PdfGState gs = new PdfGState();
		// 设置透明度为0.4
		gs.setFillOpacity(0.4f);
		gs.setStrokeOpacity(0.4f);
		
		// 设置字体
		BaseFont base = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",BaseFont.EMBEDDED);
		
		int toPage = pdfStamper.getReader().getNumberOfPages();
		
		PdfContentByte content = null;
		Rectangle pageRect = null;
		for (int i = 1; i <= toPage; i++) {
			pageRect = pdfStamper.getReader().getPageSizeWithRotation(i);
			// 计算水印X,Y坐标
			float x = pageRect.getWidth() / 2;
			float y = pageRect.getHeight() / 2;
			//获得PDF最顶层
			content = pdfStamper.getOverContent(i);
			content.saveState();
			// set Transparency
			content.setGState(gs);
			content.beginText();
			content.setColorFill(BaseColor.GRAY);
			content.setFontAndSize(base, 100);
			 // 水印文字成45度角倾斜
			content.showTextAligned(Element.ALIGN_CENTER, watermark, x, y, 315);
			content.endText();
		}
	}
}