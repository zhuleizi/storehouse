package com.cesgroup.report.config.excel.parse;

import java.util.Map;

public interface IParse {
	
	void parse(String expression, Map<String, Object> container, Object reportData);

}
