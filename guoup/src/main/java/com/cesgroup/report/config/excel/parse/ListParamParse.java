package com.cesgroup.report.config.excel.parse;

import java.util.HashMap;
import java.util.Map;

public class ListParamParse extends AbstractParse {

	@Override
	public void paramParse(String expression, Map<String, Object> container,
			Object reportData) {

		Map<String,Object> map = new HashMap<String,Object>();
		
		map.put("param", reportData);

		String[] param = expression.split(",");

        String type = param.length > 1?param[1]:null;

        Object value = getSingleValue(param[0], map,type);
        if(logger.isDebugEnabled()) {
        	//logger.debug("解析表达式：" + expression + "--->数据值：[" + value + "]");
        }
        container.put(expression, value);
	}

}
