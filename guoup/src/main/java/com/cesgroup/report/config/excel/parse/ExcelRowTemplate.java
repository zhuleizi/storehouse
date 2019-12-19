package com.cesgroup.report.config.excel.parse;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.util.CellRangeAddress;

import com.cesgroup.report.config.ExcelConfiguration;
import com.cesgroup.report.util.StringUtils;

/**
 * EXCEL行模板解析器，主要补行合并行操作
 *
 */
public class ExcelRowTemplate extends DefaultParamParse implements ExcelTemplateParse {
	
	/**
	 * 根据当前行的参数
	 */
	public int parse(HSSFSheet sheet, int currentRow, List<String> expressions, Object[] params) {
		
		// 取得数据
		List<Object> dataList = getParams(expressions, params);
		
		int returnRowNum = 0;
		int rowNum = 1;
		int subNum = 1;
		if (!dataList.isEmpty()) {
			
			for (int i = 1; i < dataList.size(); i++) {
				
				Object data = dataList.get(i);
				
				subNum = getParams(expressions, data).size();
				if (subNum > 0) {
					subNum--;
				}
				// 补行操作
				returnRowNum += insertRow(sheet, currentRow + returnRowNum + 1, rowNum + subNum, currentRow, i, data, params);
			}
			for (int i = 0; i < 1; i++) {
				Object data = dataList.get(i);
				
				subNum = getParams(expressions, data).size();
				
				if (subNum == 0) {
					subNum++;
				}
				// 补行操作
				insertRow(sheet, currentRow + 1, subNum - 1, currentRow, i, data, params);
			}
			
			// 对应下载内容高度不一致
			try {
				// for(int rowIndex = sheet.getPhysicalNumberOfRows() - 2,j =
				// sheet.getPhysicalNumberOfRows() - 1; rowIndex >= currentRow; rowIndex--) {
				// sheet.getRow(rowIndex).setHeight( sheet.getRow(j).getHeight());
				// }
				
			} catch (Exception e) {
				super.logger.warn(e.getMessage(), e);
			}
			return returnRowNum + 1;
		}
		return 0;
		
	}
	
	/**
	 * 取得当前行要补的对象
	 * 
	 * @param expressions
	 * @param data
	 * @return
	 */
	public List<Object> getParams(List<String> expressions, Object data) {
		
		for (String str : expressions) {
			
			if (str.toUpperCase().contains(ROW_LIST_MARK)) {
				
				if (str.indexOf("[].") > 0) {
					String exp = str.substring(str.indexOf(".") + 1, str.indexOf("[]."));
					Object value = this.getValue(exp, data);
					if (value instanceof List) {
						return (List<Object>) value;
					}
				}
			}
			
		}
		return new ArrayList<Object>();
	}
	
	/**
	 * 取得当前的明细对象列表
	 * 
	 * @param expressions
	 * @param params
	 * @return
	 */
	private List<Object> getParams(List<String> expressions, Object[] params) {
		for (String str : expressions) {
			if (str.toUpperCase().contains(ROW_LIST_MARK)) {
				if (str.startsWith("arg[")) {
					String ch = str.substring(4, str.indexOf("]"));
					if (ch.matches("[\\d]")) {
						int index = Integer.parseInt(ch);
						if (params[index] instanceof List) {
							return ((List) params[index]);
						}
					}
				}
			}
		}
		return new ArrayList<Object>();
	}
	
	/**
	 * 补行操作
	 * 
	 * @param sheet 对应的SHEET
	 * @param insertRow
	 * @param rowNum
	 * @param orginRowNum 当有前的行号
	 * @param rowLoopNum
	 * @param ob
	 * @return
	 */
	public int insertRow(HSSFSheet sheet, int insertRow, int rowNum, int orginRowNum, int rowLoopNum, Object ob, Object[] params) {
		
		if (rowNum > 0) {
			// 批量移动行
			sheet.shiftRows(
					insertRow, // --开端行
					sheet.getLastRowNum(), // --停止行
					rowNum, // --移动大小(行数)--往下移动
					true, // 是否复制行高
					false // 是否重置行高
			);
		}
		
		// 对批量移动后空出的空进行插入，并以插入行的上一行作为格局源(即：插入行-1的那一行)
		HSSFRow sourceRow = sheet.getRow(orginRowNum);// 源格式行
		int length = insertRow + rowNum;
		int index = insertRow;
		if (rowLoopNum == 0) {
			index--;
		}
		
		for (int i = index, j = 0; i < length; i++, j++) {
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
				
				String content = sourceCell.toString();
				
				String expression = replace(content, rowLoopNum, j);
				
				ExcelConfiguration conf = new ExcelConfiguration();
				
				Object value = conf.fillData(expression, params);
				
				conf.setCellValue(targetCell, value, expression);
				// targetCell.setCellValue(value);
				
			}
			int regions = targetRow.getSheet().getNumMergedRegions();
			for (int k = 1; k < regions; k++)// 合并单元格
			{
				CellRangeAddress cra = targetRow.getSheet().getMergedRegion(k);
				
				if (cra.getFirstRow() == sourceRow.getRowNum() && cra.getLastRow() == sourceRow.getRowNum()) {
					sheet.addMergedRegion(new CellRangeAddress(i, i, cra.getFirstColumn(), cra.getLastColumn()));
				}
			}
			
		}
		
		HSSFRow targetRow = sheet.getRow(index);
		for (int m = targetRow.getFirstCellNum(); m < targetRow.getLastCellNum(); m++)// 创建行
		{
			HSSFCell targetCell = targetRow.getCell(m);
			
			if (targetCell == null)
				continue;
			int i = index;
			for (; i < length; i++) {
				if (!StringUtils.equals(targetCell.toString(), sheet.getRow(i).getCell(m).toString())) {
					
					break;
				}
			}
			
			if ((i - 1) > index) {
				
				for (int j = index + 1; j < i; j++) {
					targetCell = sheet.getRow(j).getCell(m);
					if (targetCell == null)
						continue;
					targetCell.setCellValue("");
				}
				sheet.addMergedRegion(new CellRangeAddress(index, (i - 1), m, m));
			}
			
		}
		
		return rowNum;
	}
	
	/**
	 * 
	 * @param content
	 * @param currentRowNum
	 * @param loopNum
	 * @return
	 */
	private String replace(String content, int currentRowNum, int loopNum) {
		
		if (content.toUpperCase().contains(ROW_LIST_MARK)) {
			
			String ch = "";
			
			if (content.startsWith("arg[")) {
				ch = content.substring(4, content.indexOf("]"));
			}
			
			String header = content.substring(0, content.indexOf(".") + 1);
			
			String result = content.substring(content.indexOf(".") + 1, content.length());
			
			result = result.replaceAll("," + ROW_LIST_MARK, "");
			
			result = result.replace("[]", "[" + loopNum + "]");
			
			return header + "param" + ch + "[" + currentRowNum + "]." + result;
		}
		return "";
	}
	
}
