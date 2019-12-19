package com.cesgroup.report;

import java.util.ArrayList;
import java.util.List;

public class SheetTable {
	
	private List<SheetRef> sheetRefs = new ArrayList<SheetRef>();
	
	public SheetTable add(String cloneSheetName, List<ReportData> dataList, String... sheetNames) {
		sheetRefs.add(new SheetRef(cloneSheetName, dataList, sheetNames));
		return this;
	}
	
	public List<SheetRef> getSheetRef() {
		return sheetRefs;
	}
	
}
