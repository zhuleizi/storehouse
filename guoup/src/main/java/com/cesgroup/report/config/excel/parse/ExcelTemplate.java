package com.cesgroup.report.config.excel.parse;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.cesgroup.report.util.DateUtil;
import com.cesgroup.report.util.StringUtils;

public class ExcelTemplate extends Parse {
	
	// POI读取流
	private FileInputStream poiStream;
	// 工作薄
	private HSSFWorkbook workBook;
	
	// 当前Sheet的名称
	private String sheetName;
	
	// 工作薄解析器
	public List<ExcelTemplateParse> templateParses = new ArrayList<ExcelTemplateParse>();
	
	public ExcelTemplate() {
		
		// 增加列解析器
		templateParses.add(new ExcelColumnTemplate());
		// 增加行解析器
		templateParses.add(new ExcelRowTemplate());
		
	}
	
	/**
	 * 初始化工作薄
	 * 
	 * @param path
	 * @return
	 * @throws IOException
	 */
	public ExcelTemplate initializeWorkbook(String path) throws IOException {
		// 得到流文件
		poiStream = new FileInputStream(path);
		workBook = new HSSFWorkbook(poiStream);
		
		return this;
	}
	
	public void close() {
		try {
			poiStream.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/**
	 * 取得对应的工作SHEET
	 * 
	 * @param sheetNum
	 * @return
	 */
	public HSSFSheet getSheet(Integer sheetNum) {
		
		if (sheetNum != null) {
			return workBook.getSheetAt(sheetNum);
		} else if (StringUtils.isNotEmpty(sheetName)) {
			return workBook.getSheet(sheetName);
		} else {
			return workBook.getSheetAt(0);
		}
	}
	
	/**
	 * 设定当前Sheet的名称
	 * 
	 * @param sheetName
	 * @return
	 */
	public ExcelTemplate setCurrentSheetName(String sheetName) {
		this.sheetName = sheetName;
		return this;
	}
	
	/**
	 * 根据模板内的变量解析工作薄
	 * 
	 * @param params 填充的数据
	 * @return
	 */
	public ExcelTemplate executeTemplateParse(Object[] params) {
		
		executeTemplateParse(getSheet(null), params);
		return this;
	}
	
	/**
	 * 复制SHEET
	 * 
	 * @param clonedSheetName 要复制的SHEET名
	 * @param sheetName 生成的SHEET名
	 */
	public void cloneSheet(String clonedSheetName, String sheetName) {
		this.cloneSheet(clonedSheetName, sheetName, null);
	}
	
	/**
	 * 复制SHEET
	 * 
	 * @param clonedSheetName 要复制的SHEET名
	 * @param sheetName 生成的SHEET名
	 * @param position 生成的SHEET位置
	 */
	public void cloneSheet(String clonedSheetName, String sheetName, Integer position) {
		
		// 要复制的SHEET索引
		int sheetIndex = workBook.getSheetIndex(clonedSheetName);
		// 复制后的
		HSSFSheet sheet = this.workBook.cloneSheet(sheetIndex);
		// SHEET重命名
		workBook.setSheetName(workBook.getSheetIndex(sheet), sheetName);
		
		if (null != position) {
			// SHEET的改变索引
			workBook.setSheetOrder(sheetName, position);
		}
	}
	
	public void sortSheet(String sheetName, Integer position) {
		workBook.setSheetOrder(sheetName, position);
	}
	
	/**
	 * 重命名SHEET名
	 * 
	 * @param currentName 当前的SHEET名
	 * @param changedName 重命名后的SHEET名
	 */
	public void rename(String currentName, String changedName) {
		// SHEET重命名
		workBook.setSheetName(workBook.getSheetIndex(currentName), changedName);
	}
	
	/**
	 * 根据模板内的变量解析工作薄
	 * 
	 * @param sheet
	 * @param params
	 * @return
	 */
	public ExcelTemplate executeTemplateParse(HSSFSheet sheet, Object[] params) {
		
		List<Object> tParams = new ArrayList<Object>();
		for(int i =0;i < params.length;i++) {
			if(null != params[i]) {
				tParams.add(params[i]);
			}
		}
		params = tParams.toArray(new Object[] {});
		
		
		sheet.setForceFormulaRecalculation(true);
		//boolean isColumnMark = false;
		int rowNum = 0;
		//int executeRowNum = 0;
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		// 读取所有的商品信息
		for (int rowIndex = 0; rowIndex < physicalNumberOfRows; rowIndex++) {
			HSSFRow targetRow = sheet.getRow(rowIndex);
			if (null == targetRow)
				continue;
			List<String> expressions = getCurrentExpressions(sheet, rowIndex);
			//boolean isRowMark = false;
			if (!expressions.isEmpty()) {
				for (String expression : expressions) {
					if (expression.contains(ExcelTemplateParse.ROW_LIST_MARK)) {
						rowNum = rowIndex;
					//	isRowMark = true;
						break;
					}
					/*if (expression.contains(ExcelTemplateParse.COL_LIST_MARK)) {
						isColumnMark = true;
					}*/
				}
			}

			for (ExcelTemplateParse templateParse : templateParses) {
				templateParse.parse(sheet, rowIndex, expressions, params);
				expressions = getCurrentExpressions(sheet, rowIndex);
				for (String expression : expressions) {
					if (expression.contains(ExcelTemplateParse.ROW_LIST_MARK)) {
						rowNum = rowIndex;
					//	isRowMark = true;
						break;
					}
				}
				/*if (isRowMark && returnRowNum > 0) {
					executeRowNum = returnRowNum;
				}*/
			}
		}

	//	if (rowNum > 0) {
			
			if (params.length >= 2 && (params[1] instanceof List && !((List) params[1]).isEmpty())) {
				try {
					for (int rowIndex = sheet.getPhysicalNumberOfRows() - 2, j = sheet.getPhysicalNumberOfRows() - 1; rowIndex >= rowNum; rowIndex--) {
						sheet.getRow(rowIndex).setHeight(sheet.getRow(j).getHeight());
					}
					/*if (isColumnMark) {
						// 隐藏补样式的行
						int paddingRow = executeRowNum + rowNum;
					//	paddingColumn(sheet, paddingRow, paddingRow - 1);
						// sheet.getRow(delRow).setZeroHeight(true);
						//sheet.shiftRows(executeRowNum + rowNum , executeRowNum + rowNum , -1);
					}*/
				} catch (Exception e) {
					super.logger.warn(e.getMessage(), e);
				}
				
				//sheet.removeRow(sheet.getRow(executeRowNum + rowNum -1));
			}
	//	}
		/*if (params.length >= 2 && (params[1] instanceof List && ((List) params[1]).isEmpty())) {
			// sheet.removeRow(sheet.getRow(sheet.getPhysicalNumberOfRows() - 1));
			if (isColumnMark) {
				int paddingRow = sheet.getPhysicalNumberOfRows() - 1;
				//paddingColumn(sheet, paddingRow, paddingRow - 1);
			}
			// sheet.getRow(paddingRow).setZeroHeight(true);
		}*/
		return this;
		
	}
	

	
	/**
	 * 补列操作，并且在data行中补充相应的内容
	 * 
	 * @param sheet 对应的SHEET
	 * @param currentRowNum 标题行号
	 * @param referenceRowNum 数据行号
	 */
	private void paddingColumn(HSSFSheet sheet, int currentRowNum, int referenceRowNum) {
		
		HSSFRow currentRow = sheet.getRow(currentRowNum);
		
		HSSFRow referenceRow = sheet.getRow(referenceRowNum);
		
		if (null != currentRow && null != referenceRow) {
			
			int currentCellNums = currentRow.getPhysicalNumberOfCells();
			
			int referenceCellNums = referenceRow.getPhysicalNumberOfCells();
			
			if (currentCellNums < referenceCellNums) {
				for (int index = currentCellNums; index < referenceCellNums; index++) {
					HSSFCell cell = currentRow.createCell(index);
					cell.setCellStyle(referenceRow.getCell(index).getCellStyle());
				}
			}
		}
		
	}
	
	/**
	 * 取得当前行内的参数，变量格式${arg[0].}
	 * 
	 * @param sheet
	 * @param currentRow 当前的行号
	 * @return
	 */
	public List<String> getCurrentExpressions(HSSFSheet sheet, int currentRow) {
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
	 * 读取整个EXCEL的参数
	 * 
	 * @return
	 */
	public List<String> readExcelParams() {
		return readExcelParams(getSheet(null));
	}
	
	/**
	 * 读取整个EXCEL的参数
	 * 
	 * @param sheet
	 * @return
	 */
	public List<String> readExcelParams(HSSFSheet sheet) {
		
		List<String> params = new ArrayList<String>();
		
		// 读取所有的商品信息
		for (int rowIndex = 0; rowIndex < sheet.getLastRowNum(); rowIndex++) {
			HSSFRow row = sheet.getRow(rowIndex);
			if (null == row)
				continue;
			for (int j = 0; j < row.getLastCellNum(); j++) {
				HSSFCell cell = row.getCell(j);
				
				if (cell != null && cell.toString().contains("${")) {
					
					Matcher matcher = Pattern.compile("[$][{](.*?)[}]").matcher(cell.toString());
					while (matcher.find()) {
						String val = matcher.group();
						val = val.substring(2, matcher.group().indexOf('}'));
						params.add(val);
					}
					
				}
			}
		}
		return params;
	}
	
	/**
	 * 替换EXCEL内的变量
	 * 
	 * @param content
	 */
	public void replaceValue(Map<String, Object> content) {
		replaceValue(content, getSheet(null));
	}
	
	/**
	 * 替换EXCEL内的变量
	 * 
	 * @param content
	 * @param sheet
	 */
	public void replaceValue(Map<String, Object> content, HSSFSheet sheet) {
		
		// 读取所有的商品信息
		for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
			
			HSSFRow row = sheet.getRow(rowIndex);
			
			if (null == row)
				continue;
			// 取得每一行的列信息
			for (int j = 0; j < row.getLastCellNum(); j++) {
				
				HSSFCell cell = row.getCell(j);
				if (cell != null && cell.toString().contains("${")) {
					
					Matcher matcher = Pattern.compile("[$][{](.*?)[}]").matcher(cell.toString());
					while (matcher.find()) {
						String val = matcher.group();
						val = val.substring(2, matcher.group().indexOf('}'));
						
						String key = val;
						if (val.startsWith("arg[")) {
							key = val.substring(val.indexOf(".") + 1);
						}
						String[] params = key.split(",");
						// 单个单元格内容
						if (content.containsKey(key)) {
							Object replacement = content.get(key);
							String cellValue = cell.toString().replace("${" + val + "}", null != replacement ? replacement.toString() : "");
							setCellValue(cell, cellValue, params.length > 1 ? params[1] : null);
						} else {
							setCellValue(cell, "", params.length > 1 ? params[1] : null);
						}
						
					}
					
				}
				
			}
		}
	}
	
	/**
	 * 设定单元格的属性
	 * 
	 * @param cellValue
	 * @param type
	 */
	private void setCellValue(HSSFCell cell, String cellValue, String type) {
		type = StringUtils.isEmpty(type) ? "VARCHAR" : type.toUpperCase();
		if (StringUtils.isNotEmpty(cellValue)) {
			if (StringUtils.equals(type, "INTEGER") || StringUtils.equals(type, "DECIMAL")) {
				if (null == cellValue || StringUtils.isEmpty(cellValue)) {
					cell.setCellValue("");
				} else {
					cell.setCellValue(new BigDecimal(cellValue).doubleValue());
				}
			} else if (StringUtils.equals(type, "TIMESTAMP")) {
				cell.setCellValue(DateUtil.convertDate(cellValue));
			} else {
				cell.setCellValue(cellValue);
			}
		} else {
			if (StringUtils.equals(type, "INTEGER") || StringUtils.equals(type, "DECIMAL")) {
				String blank = null;
				cell.setCellValue(blank);
			} else {
				cell.setCellValue("");
			}
			
		}
		
	}
	
	/**
	 * 保存EXCEL文件
	 * 
	 * @param path
	 * @throws Exception
	 */
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
	
	/**
	 * 将EXCEL出力到字节流中
	 * 
	 * @return
	 * @throws Exception
	 */
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
	
	/**
	 * 根据SHEET名后缀删除相应的SHEET
	 * 
	 * @param sheetNameSuffix
	 * @throws Exception
	 */
	public void removeSheetBySuffix(String sheetNameSuffix) throws Exception {
		
		int num = workBook.getNumberOfSheets();
		List<String> names = new ArrayList<String>();
		for (int i = 0; i < num; i++) {
			String name = workBook.getSheetName(i);
			if (name.endsWith(sheetNameSuffix)) {
				names.add(name);
			}
		}
		for (String name : names) {
			workBook.removeSheetAt(workBook.getSheetIndex(name));
		}
		workBook.setActiveSheet(0);
	}
	
	public void setActiveSheet() {
		workBook.setActiveSheet(0);
	}
}
