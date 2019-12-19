package com.cesgroup.report.config.excel.parse;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;

public interface ExcelTemplateParse {

	// 行模板处理
	final String ROW_LIST_MARK = "ROW_LIST";
	
	// 列模板处理
	final String COL_LIST_MARK = "COL_LIST";
	 
	int parse(HSSFSheet sheet, int currentRow,List<String> expressions,Object[] params);
	
}
