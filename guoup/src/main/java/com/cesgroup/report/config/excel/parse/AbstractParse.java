package com.cesgroup.report.config.excel.parse;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.cesgroup.report.util.DateUtil;
import com.cesgroup.report.util.StringUtils;

import ognl.Ognl;

public abstract class AbstractParse extends Parse implements IParse {
	
	protected Log logger = LogFactory.getLog(getClass());
	
	public abstract void paramParse(String expression, Map<String, Object> container, Object reportData);
	
	/**
	 * 解析EXCEL的数据
	 * 
	 * @param expression 要解析的表达式
	 * @param container 解析后的对象
	 * @param 基础信息
	 * @param 列表信息
	 * @return
	 */
	public void parse(String expression, Map<String, Object> container, Object reportData) {
		
		// logger.debug("开始解析进行处理。");
		paramParse(expression, container, reportData);
		// logger.debug("解析进行处理结束。");
		
	}
	
	protected Object getValue(String expression, Object context) {
		try {
			return Ognl.getValue(expression, context);
		} catch (Exception e) {
			// this.logger.debug("参数[" + expression + "]从" + context + "转换有错", e);
		}
		return null;
	}
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
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
	
	/**
	 * 解析单个参数
	 * 
	 * @param expressions
	 * @param reportBaseInfo
	 * @param type
	 * @return
	 */
	protected Object getSingleValue(String expressions, Object reportBaseInfo, String type) {
		if (null == reportBaseInfo) {
			return null;
		}
		if (null != type) {
			// System.out.println(type);
		}
		Object val = getValue(expressions.split("[?]")[0], reportBaseInfo);
		return format(expressions, val, type);
		
	}
	
	/**
	 * 格式化数据
	 * 
	 * @param expressions
	 * @param val
	 * @param type
	 * @return
	 */
	protected Object format(String expressions, Object val, String type) {
		// if(null != val) {
		String[] params = expressions.split("[?]");
		if (params.length > 1) {
			if (null != type && null != val && params[1].startsWith("string(")) {
				if (StringUtils.equals("DATE", type.toUpperCase()) || StringUtils.equals("TIMESTAMP", type.toUpperCase())) {
					try {
						return DateUtil.dateToString(DateUtil.convertDate(val), params[1].substring(8, params[1].length() - 2));
					} catch (Exception e) {
						logger.warn(e.getMessage(), e);
						return null;
					}
				}
			}
			if ((null == val || StringUtils.isEmpty(val.toString())) && params[1].startsWith("default(")) {
				try {
					return params[1].substring(9, params[1].length() - 2);
				} catch (Exception e) {
					logger.warn(e.getMessage(), e);
					return null;
				}
			}
			
			// 固定值替换
			if (params[1].startsWith("replacedBy(")) {
				try {
					return params[1].substring(12, params[1].length() - 2);
				} catch (Exception e) {
					logger.warn(e.getMessage(), e);
					return null;
				}
			}
			// 单据格式
			/*
			 * if(null != val && params[1].startsWith("order") ) { return
			 * OrderUtil.getOrderId(val.toString()); }
			 */
			/**
			 * 参数1：组名 参数3: 表示名
			 */
			if (null != val && params[1].startsWith("groupKey(")) {
				String group = params[1].substring(9, params[1].length() - 1);
				try {
					String groupParms[] = group.split("[|]");
					if (groupParms.length > 0) {
						// Object ob = CodeTableUtil.getObjectByKey(groupParms[0], val.toString());
						// if (null != ob) {
						// return this.getValue(groupParms.length == 1 ? "display" : groupParms[1], ob);
						// }
					}
					
				} catch (Exception e) {
					logger.warn(e.getMessage(), e);
					return "";
				}
			}
		}
		// }
		return val;
	}
	
}
