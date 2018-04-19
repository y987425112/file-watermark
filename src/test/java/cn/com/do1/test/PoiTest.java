package cn.com.do1.test;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.sl.usermodel.TextShape.TextPlaceholder;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlideIdList;

import com.itextpdf.awt.geom.Rectangle;

public class PoiTest {
	private static Logger logger=Logger.getLogger(PoiTest.class);
	public static void main(String[] args) throws Exception {
		File fileFrom=new File("d:/m2.pptx");
		File fileTo=new File("d:/m3.pptx");
		if(fileTo.exists()) {
			fileTo.delete();
		}
		InputStream is=new FileInputStream(fileFrom);
		//HSLFSlideShow
		XMLSlideShow xmlSlideShow=new XMLSlideShow(is);
		
		List<XSLFSlide> slides = xmlSlideShow.getSlides();
		logger.info("slides.size()"+slides.size());
		XSLFSlide xslfSlide = slides.get(0);
		XSLFTextShape[] xslfTextShapes = xslfSlide.getPlaceholders();
		XSLFNotes notes = xslfSlide.getNotes();
		Iterator<XSLFShape> itShape = notes.iterator();
		while(itShape.hasNext()) {
			XSLFShape xslfShape = itShape.next();
			logger.info("xslfShape.getShapeName():"+xslfShape.getShapeName());
		}
		for(XSLFTextShape xslfTextShape:xslfTextShapes) {
			logger.info("xslfTextShape.getShapeName():"+xslfTextShape.getShapeName());
		}
		XSLFTextBox textBox = xslfSlide.createTextBox();
		textBox.setTopInset(1);
//		textBox.setText("测试水印");
//		TextPlaceholder textPlaceholder = textBox.getTextPlaceholder();
//		ShapeType shapeType = textBox.getShapeType();
//		shapeType.
		XSLFTextParagraph textParagraph=null;
		List<XSLFTextParagraph> textParagraphs = textBox.getTextParagraphs();
		logger.info("textParagraphs.size()"+textParagraphs.size());
		if(textParagraphs==null||textParagraphs.size()==0) {
			textParagraph = textBox.addNewTextParagraph();
			
		}else {
			textParagraph=textParagraphs.get(0);
		}
        List<XSLFTextRun> textRuns = textParagraph.getTextRuns();
        logger.info("textRuns.size()"+textRuns.size());
		XSLFTextRun textRun = textParagraph.addNewTextRun();
		textRun.setText("我的测试水印");
		textRun.setFontSize(10.0);
		textRun.setFontColor(Color.getHSBColor(192, 192, 192));
//		textRun.
		Dimension pageSize = xmlSlideShow.getPageSize();
		textBox.setAnchor(new Rectangle2D.Double(1, 1, 500, 10));
//		for(XSLFSlide slideTemp:slides) {
//			XSLFBackground xslfBackground = slideTemp.getBackground();
//			slideTemp.getPlaceholders();
//			xslfBackground.setPlaceholder(placeholder);
//		}
//		XSLFSlide slide = xmlSlideShow.createSlide();
//		XSLFTextBox textBox = slide.createTextBox();
////		textBox.setAnchor(new Rectangle2D.Double(10,10, 0, 0));
//		textBox.addNewTextParagraph().addNewTextRun().setText("创建幻灯片");
		OutputStream os=	new FileOutputStream(fileTo);
		xmlSlideShow.write(os);
		xmlSlideShow.close();
		is.close();
		os.close();
//		XSLFSlideShow xslfSlideShow=new XSLFSlideShow("d:/m2.pptx");
//		xslfSlideShow.
//		CTSlideIdList slideReferences = xslfSlideShow.getSlideReferences();
//		xslfSlideShow.
		
		
		
		
	
		
	}

}
