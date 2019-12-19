package com.cesgroup.report.config.excel.parse;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import ognl.Ognl;

public class Parse {
	protected Log logger = LogFactory.getLog(getClass());
	
	protected Object getValue(String expression, Object context) {
		try {
			return Ognl.getValue(expression, context);
		} catch (Exception e) {
			this.logger.debug("参数[" + expression + "]从" + context + "转换有错", e);
		}
		return null;
	}
	
	public List<String> getParameterName(String expression) {
		List names = new ArrayList();
		while (true) {
			int index = expression.indexOf("${");
			if (index < 0) {
				break;
			}
			names.add(expression.substring(index + 2, expression.indexOf("}")));
			expression = expression.substring(expression.indexOf("}") + 1, expression.length());
		}
		return names;
	}
}