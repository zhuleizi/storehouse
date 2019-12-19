package com.cesgroup.report.config.excel.parse;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import com.cesgroup.report.util.StringUtils;

public class ExcelColumnTemplate extends Parse implements ExcelTemplateParse {
	
	private boolean isPadding = false;
	
	public int parse(HSSFSheet sheet, int currentRow,
			List<String> expressions, Object[] params) {
		boolean isExecute = false;
		for (String expression : expressions) {
			// 取得数据
			List<Object> dataList = getParams(expression, params);
			if (!dataList.isEmpty()) {
				isExecute = true;
				// 补列
				if(!isPadding && dataList.size() > 1) {
					int orginCellNum = getCellIndex(sheet,currentRow,expression);
					insertCellNum( sheet, orginCellNum+1, dataList.size()-1);
				}
				addColumn(sheet, currentRow, dataList,expression);
			}
		}
		replaceColumn(sheet, currentRow);
		if(isExecute) {
			isPadding = true;
		}
		return -1;
	}

	
	private List<Object> getParams(String expressions, Object[] params) {
		if (expressions.toUpperCase().contains(COL_LIST_MARK)) {
			if (expressions.startsWith("arg[")) {
				String ch = expressions.substring(4, expressions.indexOf("]"));
				if (ch.matches("[\\d]")) {
					int index = Integer.parseInt(ch);
					Object value = this.getValue(expressions.substring(expressions.indexOf(".") + 1, expressions.indexOf("," + COL_LIST_MARK)), params[index]);
					if (value instanceof List) {
						return (List<Object>) value;
					}
				}
			}
		}
		return new ArrayList<Object>();
	}
	
	/**
	 * 补列操作，并且在data行中补充相应的内容
	 * 
	 * @param sheet 对应的SHEET
	 * @param titleRowNum 标题行号
	 * @param dataRowNum 数据行号
	 * @param titleList 需增加的标题内容
	 * @param dataModelList 需增加的数据内容
	 */
	private void addColumn(HSSFSheet sheet, int titleRowNum, List<Object> titleList,String expression) {
		// 创建HSSFRow对象
		HSSFRow titleRow = sheet.getRow(titleRowNum);
		// 取得表达式的列
		int orginCellNum = getCellIndex(sheet,titleRowNum,expression);
		if(orginCellNum < 0 ) return;
		
		HSSFCell titleCell = titleRow.getCell(orginCellNum);

		//int cellNum = titleRow.getPhysicalNumberOfCells() - 1;
		
		//orginCellNum = cellNum;
		HSSFCellStyle titleStyle = titleCell.getCellStyle();
		
		titleStyle.setBorderBottom((short) 1);
		titleStyle.setBorderLeft((short) 1);
		titleStyle.setBorderTop((short) 1);
		titleStyle.setBorderRight((short) 1);
		
		boolean isMergedRegion = isMergedRegion(sheet,titleRowNum,orginCellNum);
		
		if (null != titleList && titleList.size() > 0) {
			//String temp = "";
			for (int i = 0; i < titleList.size(); i++) {
				// 创建HSSFCell对象
				HSSFCell cell = null;
				int curCellNum = orginCellNum + i;
				
				cell = titleRow.createCell(curCellNum);
				
				cell.setCellValue(titleList.get(i).toString());
				cell.setCellStyle(titleStyle);
				sheet.setColumnWidth(curCellNum, sheet.getColumnWidth(orginCellNum));
				
				//String cellValue = titleRow.getCell(curCellNum).getStringCellValue();
				
				// CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）. 
				if (isMergedRegion && i != 0) {
					HSSFCell cell1 = sheet.getRow(titleRowNum+1).createCell(curCellNum);
					cell1.setCellStyle(titleStyle);
					sheet.addMergedRegion(new CellRangeAddress(titleRowNum, titleRowNum+1, curCellNum, curCellNum));
				}
				
				
				/*if (temp.equals(cellValue)) {
				//if (isMergedRegion(sheet,titleRowNum,curCellNum)) {
					sheet.addMergedRegion(new CellRangeAddress(titleRowNum, titleRowNum, (int) titleRow.getLastCellNum() - 2, (int) titleRow.getLastCellNum() - 1));
				}
				temp = cellValue;*/
				

			}
		}
	}
	
	/**
	 * 取得当前行
	 * @param sheet
	 * @param titleRowNum
	 * @param expression
	 * @return
	 */
	public int getCellIndex(HSSFSheet sheet, int titleRowNum,String expression) {
		StringBuffer buffer = new StringBuffer("${");
		buffer.append(expression);
		buffer.append("}");
		HSSFRow titleRow = sheet.getRow(titleRowNum);
		int cellNum = titleRow.getPhysicalNumberOfCells() - 1;
		for(int i = 0; i <= cellNum;i++) {
			HSSFCell cell = titleRow.getCell(i);
			if(null != cell) {
				if( StringUtils.eq(buffer.toString(), cell.getStringCellValue())) {
					return i;
				}
			}
		}
		return -1;
	}
	
	/**
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
    private  boolean isMergedRegion(HSSFSheet sheet, int row, int column) {
    	
        int sheetMergeCount = ((org.apache.poi.ss.usermodel.Sheet) sheet).getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = ((org.apache.poi.ss.usermodel.Sheet) sheet).getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;

    }
	
	/**
	 * 替换多余的变量
	 * 
	 * @param sheet
	 * @param currentRow
	 */
	private void replaceColumn(HSSFSheet sheet, int currentRow) {
		HSSFRow row = sheet.getRow(currentRow);
		if (null == row) {
			return;
		}
		for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
			HSSFCell cell = row.getCell(i);
			if (null != cell) {
				if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
					continue;
				}
				if (null != cell.getStringCellValue() &&
						cell.getStringCellValue().toUpperCase().contains(COL_LIST_MARK)) {
					cell.setCellValue("");
				}
			}
		}
	}
	
	/**
	 * 插入列数
	 * @param sheet
	 * @param startCell
	 * @param insertNum
	 */
	 public void insertCellNum(HSSFSheet sheet,int startCell,int insertNum){
		 int lastCellNum = 0;
		 int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		 for (int rowIndex = 0; rowIndex < physicalNumberOfRows; rowIndex++) {
			 int cellNum = sheet.getRow(rowIndex).getPhysicalNumberOfCells() - 1;
			 if(cellNum > lastCellNum) {
				 lastCellNum = cellNum;
			 }
		 }
		 // 设置列宽
		 for(int i = lastCellNum;i >= startCell;i--) {
			 sheet.setColumnWidth(i+insertNum, sheet.getColumnWidth(i));
		 }
		
		for (int rowIndex = 0; rowIndex < physicalNumberOfRows; rowIndex++) {
			HSSFRow targetRow = sheet.getRow(rowIndex);
			
			CellRangeAddress cellRangeAddress = getCellRangeAddress( sheet,  rowIndex,  0);
			
			int cellNum = targetRow.getPhysicalNumberOfCells() - 1;
			
			if(null != cellRangeAddress) {
				if(Integer.valueOf(cellRangeAddress.getLastColumn()).equals(Integer.valueOf(cellNum)) && (cellRangeAddress.getLastColumn() - cellRangeAddress.getFirstColumn()) > 0) {
					sheet.addMergedRegion(new CellRangeAddress(cellRangeAddress.getFirstRow(), cellRangeAddress.getLastRow(), 0, cellNum + insertNum));
					for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
						 CellRangeAddress range = ((org.apache.poi.ss.usermodel.Sheet) sheet).getMergedRegion(i);
						 if(range == cellRangeAddress) {
							 sheet.removeMergedRegion(i);
						 }
					}
					rowIndex = rowIndex + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow());
					continue;
				}
			}
			
			for(int index = cellNum;index >= startCell;index--) {
				cellRangeAddress = getCellRangeAddress( sheet,  rowIndex,  index);
				
				if(null == cellRangeAddress) {
					// 最后一列
					Cell originCell = targetRow.getCell(index);
					if(null != originCell) {
						Cell targetCell = targetRow.createCell(index + insertNum);
						targetCell.setCellValue(originCell.getStringCellValue());				
						targetCell.setCellStyle(originCell.getCellStyle());
					}
					targetRow.removeCell(originCell);
				} else {
						if(Integer.valueOf(cellRangeAddress.getFirstRow()).equals(Integer.valueOf(rowIndex))) {
							// 相差的列
							int subtra = cellRangeAddress.getLastColumn() -cellRangeAddress.getFirstColumn();
							index = index - subtra;
							Cell originCell = targetRow.getCell(index);
							if(null != originCell) {
								Cell targetCell = targetRow.createCell(index + insertNum);
								targetCell.setCellValue(originCell.getStringCellValue());
								targetCell.setCellStyle(originCell.getCellStyle());
								
							}
							sheet.addMergedRegion(new CellRangeAddress(cellRangeAddress.getFirstRow(), cellRangeAddress.getLastRow(),index + insertNum, index + insertNum+subtra));
							targetRow.removeCell(originCell);
							for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
								 CellRangeAddress range = ((org.apache.poi.ss.usermodel.Sheet) sheet).getMergedRegion(i);
								 if(range == cellRangeAddress) {
									 sheet.removeMergedRegion(i);
								 }
							}
						}
				}
				
			}

		}
		  
		  
	 }
	 
	    private  CellRangeAddress getCellRangeAddress(HSSFSheet sheet, int row, int column) {
	    	
	        int sheetMergeCount = ((org.apache.poi.ss.usermodel.Sheet) sheet).getNumMergedRegions();
	        for (int i = 0; i < sheetMergeCount; i++) {
	            CellRangeAddress range = ((org.apache.poi.ss.usermodel.Sheet) sheet).getMergedRegion(i);
	            int firstColumn = range.getFirstColumn();
	            int lastColumn = range.getLastColumn();
	            int firstRow = range.getFirstRow();
	            int lastRow = range.getLastRow();
	            if (row >= firstRow && row <= lastRow) {
	                if (column >= firstColumn && column <= lastColumn) {
	                    return range;
	                }
	            }
	        }
	        return null;

	    }

}
