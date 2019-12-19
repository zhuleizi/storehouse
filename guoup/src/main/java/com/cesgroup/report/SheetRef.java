package com.cesgroup.report;

import java.util.List;

public class SheetRef {
	private String cloneSheetName;
	
	private String[] sheetNames;
	
	private List<ReportData> dataList;
	
	public SheetRef(String cloneSheetName,List<ReportData> dataList,String... sheetNames) {
		this.cloneSheetName = cloneSheetName;
		this.sheetNames = sheetNames;
		this.dataList = dataList;
	}

	public String getCloneSheetName() {
		return cloneSheetName;
	}

	public void setCloneSheetName(String cloneSheetName) {
		this.cloneSheetName = cloneSheetName;
	}

	public String[] getSheetNames() {
		return sheetNames;
	}

	public void setSheetNames(String[] sheetNames) {
		this.sheetNames = sheetNames;
	}

	public List<ReportData> getDataList() {
		return dataList;
	}

	public void setDataList(List<ReportData> dataList) {
		this.dataList = dataList;
	}
}
