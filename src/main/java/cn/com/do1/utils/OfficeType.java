package cn.com.do1.utils;

/**
 * office问卷类型
 * 
 * @author ydy
 */
public enum OfficeType {
	DOC(".doc"), DOCX("docx"), XLS(".xls"), XLSX(".xlsx"), PPT(".ppt"), PPTX(".pptx");
	private OfficeType(String name) {
		this.name = name;
	}

	private String name;

	public String getName() {
		return name;
	}

}
