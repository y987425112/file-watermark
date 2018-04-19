package cn.com.do1.utils;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.log4j.Logger;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
/**
 * pdf工具
 * @author ydy
 * */

public class PdfUtil {

	private static Logger logger=Logger.getLogger(PdfUtil.class);
//	/**
//	 * 添加水印
//	 * */
//	public InputStream addWaterImage() {
//		
//		
//	}
	
	private static void addWatermark(PdfStamper pdfStamper, String waterMarkName) throws Exception {
        PdfContentByte content = null;
        BaseFont base = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",
                BaseFont.NOT_EMBEDDED);
        Rectangle pageRect = null;
        PdfGState gs = new PdfGState();
        try {
            if (base == null || pdfStamper == null) {
                return;
            }
            // 设置透明度为0.4
            gs.setFillOpacity(0.4f);
            gs.setStrokeOpacity(0.4f);
            int toPage = pdfStamper.getReader().getNumberOfPages();
            for (int i = 1; i <= toPage; i++) {
                pageRect = pdfStamper.getReader().getPageSizeWithRotation(i);
                // 计算水印X,Y坐标
                float x = pageRect.getWidth() / 2;
                float y = pageRect.getHeight() / 2;
                // 获得PDF最顶层
                content = pdfStamper.getOverContent(i);
                content.saveState();
                // set Transparency
                content.setGState(gs);
                content.beginText();
                content.setColorFill(BaseColor.GRAY);
                content.setFontAndSize(base, 30);
                // 水印文字成45度角倾斜
                content.showTextAligned(Element.ALIGN_CENTER, waterMarkName, x,
                        y, 45);
                content.endText();
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
	
	/**
	 * 添加水印
	 * @param is pdf输入流
	 * @param os pdf输出流
	 * @param waterContent 水印内容
	 * @return 是否添加成功
	 * */
	 public static  boolean addWater(InputStream is,OutputStream  os,String waterContent) {
//	        BufferedOutputStream out = new BufferedOutputStream(
//	                new FileOutputStream(new File(output)));
		 PdfReader reader=null;
		 PdfStamper stamper=null;
		 
		 try {
	        reader = new PdfReader(is);
	       
	       stamper = new PdfStamper(reader, os);
	        addWatermark(stamper,waterContent);
	        int total = reader.getNumberOfPages();
	            Image image = Image.getInstance(PdfUtil.class.getResource("/").getPath()+"/1.jpg");
	            image.setAbsolutePosition(350, 100); // set the first background image of the absolute
	            image.scaleToFit(120, 120);
	            PdfContentByte content= stamper.getOverContent(total);// 在内容上方加水印
	            content.addImage(image);
	        }catch (Exception e){
	        	e.printStackTrace();
//	        	logger.error(e.getMessage());
	        	
	        	logger.error(ExceptionUtils.getStackTrace(e));
	        	return false;
	          
	        }finally {
	        	try {
	        		if(stamper!=null) {
	        			
	        			stamper.close();
	        		}
	        		if(reader!=null) {
	        			
	        			reader.close();
	        		}
					
				} catch (Exception e2) {
					// TODO: handle exception
					logger.error(ExceptionUtils.getStackTrace(e2));
					return false;
				}
				
			}


	        return true;
	    }
	 
	 
}
