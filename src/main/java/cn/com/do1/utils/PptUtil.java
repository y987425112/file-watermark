package cn.com.do1.utils;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.UUID;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;

/**
 * ppt 工具类
 * @author ydy
 * */
public class PptUtil {
	private static Logger logger=Logger.getLogger(PptUtil.class);
	//临时目录
	private static final String TEMP_FILE_PATH;
//	private static final ActiveXComponent ppt;  
//	private static final Dispatch pptDocument ;
	
	static {
		String catalinaHome=System.getProperty("catalina.home");
		if(StringUtils.isBlank(catalinaHome)) {
			TEMP_FILE_PATH=PptUtil.class.getResource("/").getPath()+"temp";
			
		}else {//web环境
			//TEMP_FILE_PATH=catalinaHome+File.separator+"temp";
			if(catalinaHome.endsWith("/")||catalinaHome.endsWith("\\")) {
				TEMP_FILE_PATH=catalinaHome+"temp";
			}else {
				TEMP_FILE_PATH=catalinaHome+"/"+"temp";
				
			}
		}
		File fileTemp=new File(TEMP_FILE_PATH);
		if(fileTemp.exists()) {
			
		}else {
			fileTemp.mkdirs();
		}
//		ComThread.InitSTA();
//		ppt = new ActiveXComponent("PowerPoint.Application");
//		ppt.setProperty("Visible", new Variant(true));  
//		pptDocument = ppt.getProperty("Presentations").toDispatch();
//		ComThread.Release();
//		Runtime.getRuntime().addShutdownHook(new Thread(new Runnable() {
//			@Override
//			public void run()
//			{ComThread.InitSTA();
//				ppt.invoke("Quit", new Variant[] {});
//				ComThread.Release();
//			}
//		}));
		
	}
	/**
	 * ppt添加水印
	 * @param is 输入流
	 * @param os 输出流
	 * @param waterContent 水印内容
	 * @param officeType 文档类型
	 * 
	 * */
	public static synchronized  boolean addWater(InputStream is,OutputStream  os,String waterContent,OfficeType officeType) {
		try {
			//生成文件
//			String filePath =TEMP_FILE_PATH+File.separator+UUID.randomUUID().toString()+officeType.getName();
//			File fileFrom=new File(filePath);
//			fileFrom.createNewFile();
//			filePath=fileFrom.getAbsolutePath();
//			OutputStream osFileFrom=FileUtils.openOutputStream(fileFrom);
//			IOUtils.copy(is, osFileFrom);
//			osFileFrom.close();
			if(officeType==OfficeType.PPT) {//2003
				HSLFSlideShow slideShow=new HSLFSlideShow(is);
				List<HSLFSlide> slides = slideShow.getSlides();
				if(slides==null||slides.size()==0) {
					
				}else {
					for(HSLFSlide slide:slides) {
						HSLFTextBox textBox = slide.createTextBox();
						Dimension pageSize = slideShow.getPageSize();
						textBox.setAnchor(new Rectangle2D.Double(1, 1, pageSize.width-5, 10));
						List<HSLFTextParagraph> textParagraphs = textBox.getTextParagraphs();
						HSLFTextParagraph paragraph=textParagraphs.get(0);
						List<HSLFTextRun> textRuns = paragraph.getTextRuns();
						HSLFTextRun textRun = textRuns.get(0);
						textRun.setText(waterContent);
						textRun.setFontSize(10.0);
						textRun.setFontColor(Color.getHSBColor(192, 192, 192));
						
					}
				}
			}else if(officeType==officeType.PPTX) {//2007
				//ppt文件
				XMLSlideShow slideShow=new XMLSlideShow(is);
				Dimension pageSize = slideShow.getPageSize();
				List<XSLFSlide> slides = slideShow.getSlides();
				if(slides==null||slides.size()==0) {//空文档
					
				}else {
					for(XSLFSlide slide:slides) {
						//创建文本框
						XSLFTextBox textBox = slide.createTextBox();
						textBox.setAnchor(new Rectangle2D.Double(1, 1, pageSize.width-5, 10));
						XSLFTextParagraph textParagraph=null;
						List<XSLFTextParagraph> textParagraphs = textBox.getTextParagraphs();
						if(textParagraphs==null||textParagraphs.size()==0) {
							textParagraph=textBox.addNewTextParagraph();
						}else {
							textParagraph=textParagraphs.get(0);
							
						}
						XSLFTextRun textRun = textParagraph.addNewTextRun();
						textRun.setText(waterContent);
						textRun.setFontSize(10.0);
						textRun.setFontColor(Color.getHSBColor(192, 192, 192));
						
					}
				}
				slideShow.write(os);
			}else {
				return false;
			}

			return true;
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			logger.error(ExceptionUtils.getStackTrace(e));
			return false;
		}finally {
			//释放COM
			ComThread.Release();
			try {
				is.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				logger.error(ExceptionUtils.getStackTrace(e));
			}
			
			try {
				os.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				logger.error(ExceptionUtils.getStackTrace(e));
			}
		}
		
		
	}
	
	
	public static void openDocument(String filePath) {
		 //初始化com的线程
		ComThread.InitSTA();
		 ActiveXComponent ppt=null;  
//		  ActiveXComponent presentation=null;  
		      
		ppt = new ActiveXComponent("PowerPoint.Application");  
		ppt.setProperty("Visible", new Variant(true));  
		Dispatch pptDocument = ppt.getProperty("Presentations").toDispatch();
		Dispatch curDocument =Dispatch.call(pptDocument, "Open", filePath).toDispatch(); 
//Slides;
		//所有幻灯片
		Dispatch slides=	Dispatch.get(curDocument, "Slides").toDispatch();
		//获取幻灯片数量
		Variant slidesCount = Dispatch.get(slides, "Count");
		System.out.println("slidesCount:"+slidesCount);
		
		//遍历幻灯片
		for(int i=0;i<slidesCount.getInt();i++) {
			Dispatch slide= Dispatch.call(slides, "Item", new Variant(i+1)).toDispatch();
			//获取幻灯片内所有元素
		    Dispatch shapes =	Dispatch.get(slide, "Shapes").toDispatch();
		    Dispatch textEffect= Dispatch.call(shapes, "AddTextEffect",new Variant(0),"测试水印","宋体",new Variant(10),new Variant(0),new Variant(1),new Variant(0),new Variant(0)).toDispatch();
//		    textEffect
		}
		ComThread.quitMainSTA();
        
        
		
		
		
		
		
		
		
		
	}
	
	public static void main(String[] args) {
		openDocument("d:/a.pptx");
	}

}
