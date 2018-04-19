package cn.com.do1.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.junit.Test;

import com.itextpdf.text.pdf.PdfReader;

import cn.com.do1.utils.PdfUtil;

public class PdfTest {
	
	@Test
	public void test01() {
		
		//File f = new File(this.getClass().getResource("/").getPath()+"/1.jpg");
		
		try {
			InputStream is=new FileInputStream(new File("d:/5.pdf"));
			OutputStream os=new FileOutputStream(new File("d:/4.pdf"));
			
			PdfReader pdfReader=new PdfReader(is);
			System.out.println(pdfReader.isEncrypted());
			System.out.println(pdfReader.isOpenedWithFullPermissions());
			PdfUtil.addWater(is, os,"碧桂园");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		};
	}

}
