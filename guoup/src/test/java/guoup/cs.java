package guoup;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.cesgroup.report.ExcelReport;
import com.cesgroup.report.ReportData;

public class cs {
	
	public static void main(String[] args) throws Exception {
		
		ExcelReport report = new ExcelReport(false);
		
		Map<String, Object> mapT = new HashMap<String, Object>();
		mapT.put("y1", "2019");
		mapT.put("m1", "12");
		mapT.put("d1", "12");
		mapT.put("y2", "2019");
		mapT.put("m2", "12");
		mapT.put("d2", "17");
		
		List<String> list = new ArrayList<String>();
		list.add("委员证号");
		list.add("性别");
		list.add("民族");
		list.add("年龄");
		mapT.put("heads", list);
		
		list = new ArrayList<String>();
		list.add("${arg[1].number,ROW_LIST}");
		list.add("${arg[1].gender,ROW_LIST}");
		list.add("${arg[1].nation,ROW_LIST}");
		list.add("${arg[1].age,ROW_LIST}");
		mapT.put("colNames", list);
		
		List<Map> dataList = new ArrayList<Map>();
		Map<String, String> map = new HashMap<String, String>();
		map.put("name", "张三");
		map.put("number", "001");
		map.put("gender", "男");
		// map.put("nation", "汉族");
		// map.put("age", "30");
		map.put("cccc", "李四");
		dataList.add(map);
		
		map = new HashMap<String, String>();
		map.put("name", "李四");
		map.put("cccc", "李四1");
		map.put("number", "001");
		dataList.add(map);
		
		report.exportExcel(new ReportData(mapT, dataList), "E:/yhryxx.xls", "E://test//");
		
	}
	
}
