package cn.com.do1.utils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.UUID;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.log4j.Logger;

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
		ComThread.InitSTA();
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
			String filePath =TEMP_FILE_PATH+File.separator+UUID.randomUUID().toString()+officeType.getName();
			File fileFrom=new File(filePath);
			fileFrom.createNewFile();
			filePath=fileFrom.getAbsolutePath();
			OutputStream osFileFrom=FileUtils.openOutputStream(fileFrom);
			IOUtils.copy(is, osFileFrom);
			
			osFileFrom.close();
			//初始化com的线程
		
			ComThread.InitSTA();
			//ppt程序
			ActiveXComponent ppt = new ActiveXComponent("PowerPoint.Application");
			Dispatch pptDocument = ppt.getProperty("Presentations").toDispatch();
			//PasswordEncryptionFileProperties 
			
			ppt.setProperty("Visible", new Variant(true));  
			
			//打开文档
			Dispatch curDocument =Dispatch.call(pptDocument, "Open", filePath,false,false,false).toDispatch();
			Variant isPassword= Dispatch.get(curDocument, "PasswordEncryptionFileProperties");
//			Variant isPassword= Dispatch.get(curDocument, "Password");
			System.out.println("isPassword:"+isPassword.toString());
			//所有幻灯片
			Dispatch slides=	Dispatch.get(curDocument, "Slides").toDispatch();
			//获取幻灯片数量
			Variant slidesCount = Dispatch.get(slides, "Count");
			
			//遍历幻灯片
			for(int i=0;i<slidesCount.getInt();i++) {
				Dispatch slide= Dispatch.call(slides, "Item", new Variant(i+1)).toDispatch();
				//获取幻灯片内所有元素
			    Dispatch shapes =	Dispatch.get(slide, "Shapes").toDispatch();
			    //添加水印
			    Dispatch.call(shapes, "AddTextEffect",new Variant(0),waterContent,"宋体",new Variant(10),new Variant(0),new Variant(1),new Variant(0),new Variant(0)).toDispatch();
			}
			String filePathTo =TEMP_FILE_PATH+File.separator+UUID.randomUUID().toString()+officeType.getName();
			File fileTo=new File(filePathTo);
			//保存
			Dispatch.call(curDocument, "SaveAs",fileTo.getAbsolutePath());
			//关闭文件
			Dispatch.call(curDocument, "Close");
			//关闭程序
//			ppt.invoke("Quit", new Variant[] {});
			
			InputStream isFileTO=FileUtils.openInputStream(fileTo);
			IOUtils.copy(isFileTO, os);
			
			isFileTO.close();
			return true;
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			logger.error(ExceptionUtils.getStackTrace(e));
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
		return false;
		
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
