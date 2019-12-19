package com.cesgroup.report.config.excel.parse;

import java.util.List;
import java.util.Map;

/**
 * 默认值解析器
 */
public class DefaultParamParse extends AbstractParse {
	
	@SuppressWarnings("rawtypes")
	public void paramParse(String expression, Map<String, Object> container,
			Object reportData) {
		
		String[] param = expression.split(",");
		// 取得数据类型
		String type = param.length > 1 ? param[1] : null;
		Object value = null;
		if (null == reportData ||
				(reportData instanceof List && ((null == (List) reportData) || ((List) reportData).isEmpty()))) {
			value = null;
		} else {
			value = getSingleValue(param[0], reportData, type);
		}
		
		/*
		 * if(logger.isDebugEnabled()){ logger.debug("解析表达式：" + expression + "--->数据值：["
		 * + value + "]"); }
		 */
		
		container.put(expression, value);
		
	}
	
	public String parse(String expression, Object reportData) {
		String[] param = expression.split(",");
		// 取得数据类型
		String type = param.length > 1 ? param[1] : null;
		Object value = null;
		if (null == reportData ||
				(reportData instanceof List && ((null == (List) reportData) || ((List) reportData).isEmpty()))) {
			value = null;
		} else {
			value = getSingleValue(param[0], reportData, type);
		}
		return null != value ? value.toString() : "";
	}
	
}
