package com.cesgroup.report;

import java.util.ArrayList;
import java.util.List;

/**
 * 报表数据
 * 
 * @author Sakamoto
 */
public class ReportData {
	private Object baseInfo = new Object(); // 基础数据
	private List<?> datalist = new ArrayList(); // 多行数据
	
	public ReportData() {
		
	}
	
	public ReportData(Object baseInfo, List<?> datalist) {
		this.baseInfo = baseInfo;
		this.datalist = datalist;
	}
	
	public Object getBaseInfo() {
		return baseInfo;
	}
	
	public void setBaseInfo(Object baseInfo) {
		this.baseInfo = baseInfo;
	}
	
	public List<?> getDatalist() {
		return datalist;
	}
	
	public void setDatalist(List<?> datalist) {
		this.datalist = datalist;
	}
	
}
