package com.cesgroup.report.config;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;

import com.cesgroup.report.config.excel.ParamParseMaker;
import com.cesgroup.report.util.DateUtil;
import com.cesgroup.report.util.StringUtils;

/**
 * 
 * EXCEL 解析配置器
 * 
 * @author Sakamoto
 *
 */
public class ExcelConfiguration {
	
	// EXCEL自定义参数解析工厂
	private ParamParseMaker maker;
	
	public ExcelConfiguration() {
		maker = ParamParseMaker.getInstance();
	}
	
	/**
	 * 填充参数
	 * 
	 * @param expressions
	 * @param reportBaseInfo
	 * @param reportDataList
	 * @return
	 * 
	 * @author Sakamoto
	 */
	public Map<String, Object> fillData(List<String> expressions, Object[] reportData) {
		// 数据字典容器
		Map<String, Object> container = new HashMap<String, Object>();
		
		for (String expression : expressions) {
			if (expression.startsWith("arg")) {
				String ch = expression.substring(4, expression.indexOf(".") - 1);
				if (ch.matches("[\\d]")) {
					int index = Integer.parseInt(ch);
					expression = expression.substring(expression.indexOf(".") + 1, expression.length());
					maker.createParse(expression).parse(expression, container, reportData[index]);
				}
				
			}
			
		}
		
		return container;
	}
	
	public Object fillData(String expression, Object[] reportData) {
		
		if (expression.length() < 2) {
			return "";
		}
		
		expression = expression.substring(2, expression.length() - 1);
		
		Map<String, Object> container = new HashMap<String, Object>();
		if (expression.startsWith("arg")) {
			String ch = expression.substring(4, expression.indexOf(".") - 1);
			if (ch.matches("[\\d]")) {
				int index = Integer.parseInt(ch);
				expression = expression.substring(expression.indexOf(".") + 1, expression.length());
				maker.createParse(expression).parse(expression, container, reportData[index]);
			}
			
		}
		return container.get(expression);
	}
	
	public void setCellValue(HSSFCell cell, Object cellValue, String expression) {
		if (expression.length() < 2) {
			cell.setCellValue("");
		} else {
			
			String val = expression.substring(2, expression.length() - 1);
			
			String key = val;
			if (val.startsWith("arg[")) {
				key = val.substring(val.indexOf(".") + 1);
			}
			
			String[] params = key.split(",");
			
			String type = params.length > 1 ? params[1] : null;
			// 单个单元格内容
			type = StringUtils.isEmpty(type) ? "VARCHAR" : type.toUpperCase();
			if (null != cellValue) {
				
				if (StringUtils.isNotEmpty(cellValue.toString())) {
					if (StringUtils.equals(type, "INTEGER") || StringUtils.equals(type, "DECIMAL")) {
						if (null == cellValue || StringUtils.isEmpty(cellValue.toString())) {
							cell.setCellValue("");
						} else {
							cell.setCellValue(new BigDecimal(cellValue.toString()).doubleValue());
						}
					} else if (StringUtils.equals(type, "TIMESTAMP")) {
						cell.setCellValue(DateUtil.convertDate(cellValue));
					} else {
						cell.setCellValue(cellValue.toString());
					}
				} else {
					if (StringUtils.equals(type, "INTEGER") || StringUtils.equals(type, "DECIMAL")) {
						String blank = null;
						cell.setCellValue(blank);
					} else {
						cell.setCellValue("");
					}
					
				}
			} else {
				cell.setCellValue("");
			}
		}
	}
	
}
