package com.cesgroup.report.util;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * POI报表共通处理类
 * 
 * @author Sakamoto
 */
public class ExcelUtil {
	
	// POI读取流
	private FileInputStream poiStream;
	//
	private HSSFWorkbook workBook;
	
	public ExcelUtil initializeWorkbook(String path) throws IOException {
		// 得到流文件
		poiStream = new FileInputStream(path);
		workBook = new HSSFWorkbook(poiStream);
		return this;
	}
	
	public void close() throws IOException {
		if (null != poiStream) {
			poiStream.close();
		}
	}
	
	public static List<String> getCurrentExpressions(HSSFSheet sheet, int currentRow) {
		// 当前的行表达式
		List<String> expressions = new ArrayList<String>();
		HSSFRow row = sheet.getRow(currentRow);
		for (int j = 0; j < row.getLastCellNum(); j++) {
			HSSFCell cell = row.getCell(j);
			
			if (cell != null && cell.toString().contains("${")) {
				
				Matcher matcher = Pattern.compile("[$][{](.*?)[}]").matcher(cell.toString());
				while (matcher.find()) {
					String val = matcher.group();
					val = val.substring(2, matcher.group().indexOf('}'));
					expressions.add(val);
				}
				
			}
			
		}
		return expressions;
		
	}
	
	/**
	 * 读取excel模板文件
	 * 
	 * @param inputStream 获取文件目录
	 * @return
	 * @throws Exception
	 * @throws IOException
	 * 
	 */
	public Map<String, String> readExcelParams() {
		Map<String, String> params = new HashMap<String, String>();
		// 获取第一个sheet
		HSSFSheet sheet = workBook.getSheetAt(0);
		// 读取所有的商品信息
		for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
			HSSFRow row = sheet.getRow(rowIndex);
			for (int j = 0; j < row.getLastCellNum(); j++) {
				HSSFCell cell = row.getCell(j);
				
				if (cell != null && cell.toString().contains("${")) {
					
					Matcher matcher = Pattern.compile("[$][{](.*?)[}]").matcher(cell.toString());
					while (matcher.find()) {
						String val = matcher.group();
						val = val.substring(2, matcher.group().indexOf('}'));
						params.put(val, val);
					}
					
				}
				
			}
		}
		
		return params;
	}
	
	/**
	 * 设置单元格
	 * 
	 * @param cell
	 * @return
	 */
	public String settingCell(HSSFCell cell) {
		String cellValue = "";
		if (cell != null) {
			// 判断当前Cell的Type
			switch (cell.getCellType()) {
				// 如果当前Cell的Type为NUMERIC
				case HSSFCell.CELL_TYPE_NUMERIC: {
					BigDecimal big = new BigDecimal(cell.getNumericCellValue());
					cellValue = big.toString();
				}
				break;
				case HSSFCell.CELL_TYPE_FORMULA: {
					// 判断当前的cell是否为Date
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						// data->string：yyyy-MM-dd HH:mm
						cellValue = DateUtil.formatDate2String(cell.getDateCellValue(), DateUtil.FORMAT_TYPE_MINUTE);
					}
					// 如果是纯数字
					else {
						// 取得当前Cell的数值
						cellValue = String.valueOf(cell.getNumericCellValue());
					}
				}
				break;
				// 如果当前Cell的Type为string
				case HSSFCell.CELL_TYPE_STRING:
					// 取得当前的Cell字符串
					cellValue = cell.getRichStringCellValue().getString();
				break;
				// 默认的Cell值
				default:
					cellValue = "";
			}
		} else {
			cellValue = "";
		}
		return cellValue;
	}
	
	/**
	 * 替换excel里的值
	 * 
	 * @param content
	 */
	@SuppressWarnings("unchecked")
	public void replaceValue(Map<String, Object> content) {
		
		List<Integer> rowNums = new ArrayList<Integer>();
		
		List<String> rowRelationKeys = new ArrayList<String>();
		
		// 获取第一个sheet
		HSSFSheet sheet = workBook.getSheetAt(0);
		
		// 读取所有的商品信息
		for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
			
			boolean isFirstCell = true;
			
			HSSFRow row = sheet.getRow(rowIndex);
			// 取得每一行的列信息
			for (int j = 0; j < row.getLastCellNum(); j++) {
				
				HSSFCell cell = row.getCell(j);
				if (cell != null && cell.toString().contains("${")) {
					
					Matcher matcher = Pattern.compile("[$][{](.*?)[}]").matcher(cell.toString());
					while (matcher.find()) {
						String val = matcher.group();
						val = val.substring(2, matcher.group().indexOf('}'));
						// 如果参数中包含LIST信息
						if (val.toString().toLowerCase().endsWith("list")) {
							if (isFirstCell) {
								rowNums.add(row.getRowNum());
								rowRelationKeys.add(val);
								isFirstCell = false;
							}
						} else {
							// 单个单元格内容
							if (content.containsKey(val)) {
								Object replacement = content.get(val);
								String cellValue = cell.toString().replace("${" + val + "}", null != replacement ? replacement.toString() : "");
								cell.setCellValue(cellValue);
							}
						}
						
					}
					
				}
				
			}
		}
		
		// 多条数据集
		int returnRowNum = 0;
		for (int i = 0; i < rowNums.size(); i++) {
			String[] key = rowRelationKeys.get(i).split(",");
			List<Map<String, Object>> dataList = (List<Map<String, Object>>) content.get(key[0] + "_" + key[2]);
			if (null != dataList && dataList.size() > 0)
				returnRowNum += InsertRow(sheet, rowNums.get(i) + returnRowNum + 1, dataList);
		}
		
	}
	
	public int InsertRow(HSSFSheet sheet, int insertRow, List<Map<String, Object>> dataList) {
		// 要插入的行数
		int rownum = dataList.size() - 1;
		
		// 批量移动行
		sheet.shiftRows(
				insertRow, // --开端行
				sheet.getLastRowNum(), // --停止行
				rownum, // --移动大小(行数)--往下移动
				true, // 是否复制行高
				false // 是否重置行高
		);
		// 对批量移动后空出的空进行插入，并以插入行的上一行作为格局源(即：插入行-1的那一行)
		HSSFRow sourceRow = sheet.getRow(insertRow - 1);// 源格式行
		
		if (dataList.size() > 1) {
			
			for (int i = insertRow, j = 0; i < insertRow + rownum; i++, j++) {
				HSSFRow targetRow = sheet.createRow(i);// 目标行
				HSSFCell sourceCell = null;// 源行单元格
				HSSFCell targetCell = null;// 目标行单元格
				for (int m = sourceRow.getFirstCellNum(); m < sourceRow.getLastCellNum(); m++)// 创建行
				{
					sourceCell = sourceRow.getCell(m);
					if (sourceCell == null)
						continue;
					targetCell = targetRow.createCell(m);
					
					targetCell.setCellStyle(sourceCell.getCellStyle());
					
					targetCell.setCellType(sourceCell.getCellType());
					
					targetCell.getRow().setHeight(sourceCell.getRow().getHeight());
					
					Map<String, Object> dic = dataList.get(j + 1);
					
					String content = sourceCell.toString();
					for (String str : dic.keySet()) {
						if (content.toString().contains(str)) {
							content = content.replace("${" + str + "}", dic.get(str).toString());
						}
					}
					targetCell.setCellValue(content);
					
				}
				
				for (int k = 1; k < targetRow.getSheet().getNumMergedRegions(); k++)// 合并单元格
				{
					CellRangeAddress cra = targetRow.getSheet().getMergedRegion(k);
					if (cra.getFirstRow() == sourceRow.getRowNum() && cra.getLastRow() == sourceRow.getRowNum()) {
						sheet.addMergedRegion(new CellRangeAddress(i, i, cra.getFirstColumn(), cra.getLastColumn()));
					}
				}
			}
		}
		// 原数据
		for (int m = sourceRow.getFirstCellNum(); m < sourceRow.getLastCellNum(); m++)// 创建行
		{
			HSSFCell sourceCell = sourceRow.getCell(m);
			if (sourceCell == null)
				continue;
			Map<String, Object> dic = dataList.get(0);
			
			String content = sourceCell.toString();
			for (String str : dic.keySet()) {
				if (content.toString().contains(str)) {
					content = content.replace("${" + str + "}", dic.get(str).toString());
				}
			}
			sourceCell.setCellValue(content);
		}
		
		return rownum;
	}
	
	public static int insertRow(HSSFSheet sheet, int insertRow, int rowNum, Object ob) {
		
		// 批量移动行
		sheet.shiftRows(
				insertRow, // --开端行
				sheet.getLastRowNum(), // --停止行
				rowNum, // --移动大小(行数)--往下移动
				true, // 是否复制行高
				false // 是否重置行高
		);
		
		// 对批量移动后空出的空进行插入，并以插入行的上一行作为格局源(即：插入行-1的那一行)
		HSSFRow sourceRow = sheet.getRow(insertRow - 1);// 源格式行
		
		for (int i = insertRow, j = 0; i < insertRow + rowNum; i++, j++) {
			HSSFRow targetRow = sheet.createRow(i);// 目标行
			HSSFCell sourceCell = null;// 源行单元格
			HSSFCell targetCell = null;// 目标行单元格
			for (int m = sourceRow.getFirstCellNum(); m < sourceRow.getLastCellNum(); m++)// 创建行
			{
				sourceCell = sourceRow.getCell(m);
				if (sourceCell == null)
					continue;
				targetCell = targetRow.createCell(m);
				
				targetCell.setCellStyle(sourceCell.getCellStyle());
				
				targetCell.setCellType(sourceCell.getCellType());
				
				targetCell.getRow().setHeight(sourceCell.getRow().getHeight());
				
				//
				String content = sourceCell.toString();
				
				targetCell.setCellValue(content);
				
			}
			
			for (int k = 1; k < targetRow.getSheet().getNumMergedRegions(); k++)// 合并单元格
			{
				CellRangeAddress cra = targetRow.getSheet().getMergedRegion(k);
				
				if (cra.getFirstRow() == sourceRow.getRowNum() && cra.getLastRow() == sourceRow.getRowNum()) {
					sheet.addMergedRegion(new CellRangeAddress(i, i, cra.getFirstColumn(), cra.getLastColumn()));
				}
			}
			
		}
		
		return rowNum;
	}
	
	public void saveToFile(String path) throws Exception {
		FileOutputStream fstream = new FileOutputStream(path);
		try {
			workBook.write(fstream);
			
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			fstream.close();
		}
		
	}
	
	public byte[] toBytes() throws Exception {
		ByteArrayOutputStream arrayOutputStream = new ByteArrayOutputStream();
		try {
			workBook.write(arrayOutputStream);
			byte[] bytes = arrayOutputStream.toByteArray();
			return bytes;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			arrayOutputStream.close();
		}
		return new byte[0];
	}
	
	public static void main2(String[] args) throws IOException {
		
		ExcelUtil testT = new ExcelUtil();
		Map<String, String> map = testT.initializeWorkbook("D:\\207污染排放季报.xls").readExcelParams();
		
		// System.out.print(map);
		// System.out.println("获得Excel表格的内容:");
		for (int i = 1; i <= map.size(); i++) {
			// System.out.println(map.get(i));
		}
	}
	
	/**
	 * @param args the command line arguments
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		// TODO code application logic here
		FileInputStream myxls = new FileInputStream("D:\\207污染排放季报.xls");
		HSSFWorkbook wb = new HSSFWorkbook(myxls);
		HSSFSheet sheet = wb.getSheetAt(0);
		int startRow = 9;
		int rows = 10;
		
		insertRow(sheet, startRow, rows);
		
		FileOutputStream myxlsout = new FileOutputStream("D:\\workbook.xls");
		wb.write(myxlsout);
		myxlsout.close();
		
	}
	
	private static void insertRow(HSSFSheet sheet, int startRow, int rows) {
		
		sheet.shiftRows(startRow, sheet.getLastRowNum(), rows, true, false);
		
		// 对批量移动后空出的空进行插入，并以插入行的上一行作为格局源(即：插入行-1的那一行)
		HSSFRow sourceRow = sheet.getRow(startRow - 1);// 源格式行
		
		for (int i = startRow, j = 0; i < startRow + rows; i++, j++) {
			HSSFRow targetRow = sheet.createRow(i);// 目标行
			HSSFCell sourceCell = null;// 源行单元格
			HSSFCell targetCell = null;// 目标行单元格
			for (int m = sourceRow.getFirstCellNum(); m < sourceRow.getLastCellNum(); m++)// 创建行
			{
				sourceCell = sourceRow.getCell(m);
				if (sourceCell == null)
					continue;
				targetCell = targetRow.createCell(m);
				
				targetCell.setCellStyle(sourceCell.getCellStyle());
				
				targetCell.setCellType(sourceCell.getCellType());
				
				targetCell.getRow().setHeight(sourceCell.getRow().getHeight());
				
			}
			
			for (int k = 1; k < targetRow.getSheet().getNumMergedRegions(); k++)// 合并单元格
			{
				CellRangeAddress cra = targetRow.getSheet().getMergedRegion(k);
				if (cra.getFirstRow() == sourceRow.getRowNum() && cra.getLastRow() == sourceRow.getRowNum()) {
					sheet.addMergedRegion(new CellRangeAddress(i, i, cra.getFirstColumn(), cra.getLastColumn()));
				}
			}
		}
		
	}
	
}
