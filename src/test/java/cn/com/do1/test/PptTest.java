package cn.com.do1.test;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.commons.io.FileUtils;
import org.junit.Test;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;

import cn.com.do1.utils.OfficeType;
import cn.com.do1.utils.PptUtil;

public class PptTest {
//	@Test
//	public void test() {
//		try {
//			//并发测试
//			File fileTempate=new File("d:/1.pptx");
//			for(int i=0;i<10;i++) {
//				String filePath="d:/a"+i+".pptx";
//				File fileTemp=new  File(filePath);
//				if(fileTemp.exists()) {
//					fileTemp.delete();
//				}
//				fileTemp.createNewFile();
//				FileUtils.copyFile(fileTempate, fileTemp);
//			}
//			for(int i=0;i<10;i++) {
//				final int y=i;
//				Runnable runnable=new Runnable() {
//					
//					@Override
//					public void run() {
//						// TODO Auto-generated method stub
//						try {
//							
//							File inFile=new File("d:/a"+y+".pptx");
//							File outFile=new File("d:/b"+y+".pptx");
//							outFile.deleteOnExit();
//							InputStream is=FileUtils.openInputStream(inFile);
//							OutputStream os=FileUtils.openOutputStream(outFile);
//							PptUtil.addWater(is, os, "东营浪人测试水印", OfficeType.PPTX);
//						} catch (Exception e) {
//							// TODO: handle exception
//							e.printStackTrace();
//						}
//						
//						
//						
//					}
//				};
//				Thread t=new Thread(runnable);
//				t.start();
//			}
//			
////			File inFile=new File("d:/1.pptx");
////			File outFile=new File("d:/2.pptx");
////			if(outFile.exists()) {
////				outFile.delete();
////			}
////			outFile.createNewFile();
////			InputStream is=FileUtils.openInputStream(inFile);
////			OutputStream os=FileUtils.openOutputStream(outFile);
////			PptUtil.addWater(is, os, "东营浪人测试水印", OfficeType.PPTX);
//		} catch (Exception e) {
//			// TODO: handle exception
//			e.printStackTrace();
//		}
//		
//	}
	
//	public static void main(String[] args) {
//		try {
//		long startTime=	System.currentTimeMillis();
//			//并发测试
//			File fileTempate=new File("d:/1.pptx");
//			int count=1;
//			for(int i=0;i<count;i++) {
//				String filePath="d:/a"+i+".pptx";
//				File fileTemp=new  File(filePath);
//				if(fileTemp.exists()) {
//					fileTemp.delete();
//				}
//				fileTemp.createNewFile();
//				FileUtils.copyFile(fileTempate, fileTemp);
//			}
//			for(int i=0;i<count;i++) {
//				final int y=i;
//				Runnable runnable=new Runnable() {
//					
//					@Override
//					public void run() {
//						// TODO Auto-generated method stub
//						try {
//							
//							File inFile=new File("d:/a"+y+".pptx");
//							File outFile=new File("d:/b"+y+".pptx");
//							if(outFile.exists()) {
//								outFile.delete();
//							}
//							InputStream is=FileUtils.openInputStream(inFile);
//							OutputStream os=FileUtils.openOutputStream(outFile);
//							PptUtil.addWater(is, os, "东营浪人测试水印", OfficeType.PPTX);
//						} catch (Exception e) {
//							// TODO: handle exception
//							e.printStackTrace();
//						}
//						
//						
//						
//					}
//				};
////				Thread t=new Thread(runnable);
////				t.start();
//				runnable.run();
//			}
//			long endTime=	System.currentTimeMillis();
//			long timeDiff=endTime-startTime;
//			System.out.println("timeDiff:"+timeDiff/1000+"s");
////			File inFile=new File("d:/1.pptx");
////			File outFile=new File("d:/2.pptx");
////			if(outFile.exists()) {
////				outFile.delete();
////			}
////			outFile.createNewFile();
////			InputStream is=FileUtils.openInputStream(inFile);
////			OutputStream os=FileUtils.openOutputStream(outFile);
////			PptUtil.addWater(is, os, "东营浪人测试水印", OfficeType.PPTX);
//		} catch (Exception e) {
//			// TODO: handle exception
//			e.printStackTrace();
//		}
//	}
	@Test
	public void test2() {
	try {
		
		File inFile=new File("d:/m2.pptx");
		File outFile=new File("d:/99.pptx");
		if(outFile.exists()) {
			outFile.delete();
		}
		outFile.createNewFile();
		InputStream is=FileUtils.openInputStream(inFile);
		OutputStream os=FileUtils.openOutputStream(outFile);
		PptUtil.addWater(is, os, "东营浪人测试水印", OfficeType.PPTX);
	} catch (Exception e) {
		// TODO: handle exception
		e.printStackTrace();
	}	
		
	}

}
