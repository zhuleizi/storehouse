package com.cesgroup.report;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.channels.FileChannel.MapMode;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.commons.io.FileUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.cesgroup.report.config.ExcelConfiguration;
import com.cesgroup.report.config.excel.parse.ExcelTemplate;
import com.cesgroup.report.util.DateUtil;
import com.cesgroup.report.util.StringUtils;

/**
 * @author Sakamoto
 */
public class ExcelReportBytes {
	
	private ExcelTemplate excelTemplate = new ExcelTemplate();
	
	protected Log logger = LogFactory.getLog(getClass());
	
	private boolean isDelTemp = false;
	
	public ExcelReportBytes() {
		this(true);
	}
	
	public ExcelReportBytes(boolean isDelTemp) {
		this.isDelTemp = isDelTemp;
	}
	
	/**
	 * 根据EXCEL模板导入EXCEL文件（单sheet）
	 * 
	 * @param reportDatas
	 * @param templatePath
	 * @param tempDictionaryPath
	 * @return
	 * @throws Exception
	 * @author Sakamoto
	 */
	public byte[] exportExcel(ReportData reportDatas, String templatePath, String tempDictionaryPath) throws Exception {
		return this.exportExcel(new Object[] { reportDatas.getBaseInfo(), reportDatas.getDatalist() }, templatePath, tempDictionaryPath);
	}
	
	/**
	 * 根据EXCEL模板导入EXCEL文件（单sheet）
	 * 
	 * @param Object 报表数据信息
	 * @param templatePath 模板的路径
	 * @param tempDictionaryPath 临时目录的路径
	 * @return
	 * @throws Exception
	 * @author Sakamoto
	 */
	public byte[] exportExcel(Object[] reportDatas, String templatePath, String tempDictionaryPath) throws Exception {
		byte[] bytes = new byte[0];
		
		// 临时文件
		String tmpFileName = createTempTemplateFile(templatePath, tempDictionaryPath);
		
		logger.debug("打开EXCEL临时的模板文件PATH=" + tmpFileName);
		// 理由：没有数据的场合
		/*
		 * boolean isEmptyData = false; try { if (null != reportDatas &&
		 * reportDatas.length == 2 && reportDatas[1] instanceof java.util.ArrayList) {
		 * if (null == ((java.util.ArrayList) reportDatas[1]) || ((java.util.ArrayList)
		 * reportDatas[1]).isEmpty()) { isEmptyData = true; } } } catch (Exception e) {
		 * logger.warn(e.getMessage(), e); }
		 */
		
		excelTemplate.initializeWorkbook(tmpFileName).executeTemplateParse(reportDatas);
		
		excelTemplate.saveToFile(tmpFileName);
		
		// 初始化工作薄取得相应的参数
		// List<String> excelParams = excelTemplate.readExcelParams();
		
		// ExcelConfiguration conf = new ExcelConfiguration();
		
		// Map<String, Object> content = conf.fillData(excelParams, reportDatas);
		// 理由：没有数据的场合
		// base信息需要解析头信息
		// if(isEmptyData){
		List<String> excelParams = excelTemplate.readExcelParams();
		
		ExcelConfiguration conf = new ExcelConfiguration();
		Map<String, Object> content = conf.fillData(excelParams, reportDatas);
		excelTemplate.replaceValue(content);
		// }
		
		logger.debug("替换模板数据。");
		
		// excelTemplate.replaceValue(content);
		
		excelTemplate.saveToFile(tmpFileName);
		
		logger.debug("临时文件保存成功PATH=" + tmpFileName);
		
		bytes = excelTemplate.toBytes();
		
		if (isDelTemp) {
			new File(tmpFileName).delete();
			logger.debug("临时文件删除成功PATH=" + tmpFileName);
		}
		
		return bytes;
	}
	
	/**
	 * 根据EXCEL模板导入EXCEL文件（多sheet）
	 * 
	 * @param sheetRef 数据
	 * @param templateFile 模板文件
	 * @param tempDir 临时目录
	 * @return
	 * @throws Exception
	 * @author Sakamoto
	 */
	public byte[] exportExcel(SheetRef sheetRef, String templateFile, String tempDir) throws Exception {
		
		byte[] bytes = new byte[0];
		
		String tmpFileName = createTempTemplateFile(templateFile, tempDir);
		
		for (int index = 0; index < sheetRef.getSheetNames().length; index++) {
			if (StringUtils.isNotEmpty(sheetRef.getCloneSheetName())) {
				cloneTemplateSheet(tmpFileName, sheetRef.getCloneSheetName(), sheetRef.getSheetNames()[index], Integer.valueOf(0));
			}
			
			parseWorkBookBySheetName(sheetRef.getDataList().get(index), tmpFileName, sheetRef.getSheetNames()[index]);
		}
		
		if (StringUtils.isNotEmpty(sheetRef.getCloneSheetName())) {
			removeSheetBySuffix(tmpFileName, sheetRef.getCloneSheetName());
		} else {
			excelTemplate.setActiveSheet();
		}
		saveToFile(tmpFileName);
		
		bytes = getBytes(tmpFileName);
		if (isDelTemp) {
			new File(tmpFileName).delete();
			logger.debug("临时文件删除成功PATH=" + tmpFileName);
		}
		
		return bytes;
	}
	
	/**
	 * 根据EXCEL模板导入EXCEL文件（多sheet）
	 * 
	 * @param sheetTable
	 * @param reportDatas
	 * @param templateFile
	 * @param tempDir
	 * @return
	 * @throws Exception
	 * @author Sakamoto
	 */
	public byte[] exportExcel(SheetTable sheetTable, String templateFile, String tempDir) throws Exception {
		
		byte[] bytes = new byte[0];
		// 生成临时文件
		String tmpFileName = createTempTemplateFile(templateFile, tempDir);
		
		List<SheetRef> sheetRefs = sheetTable.getSheetRef();
		
		for (SheetRef sheetRef : sheetRefs) {
			
			for (int index = 0; index < sheetRef.getSheetNames().length; index++) {
				if (StringUtils.isNotEmpty(sheetRef.getCloneSheetName())) {
					cloneTemplateSheet(tmpFileName, sheetRef.getCloneSheetName(), sheetRef.getSheetNames()[index], Integer.valueOf(0));
				}
				
				parseWorkBookBySheetName(sheetRef.getDataList().get(index), tmpFileName, sheetRef.getSheetNames()[index]);
			}
			
			if (StringUtils.isNotEmpty(sheetRef.getCloneSheetName())) {
				removeSheetBySuffix(tmpFileName, sheetRef.getCloneSheetName());
			}
		}
		saveToFile(tmpFileName);
		
		bytes = getBytes(tmpFileName);
		
		if (isDelTemp) {
			new File(tmpFileName).delete();
			logger.debug("临时文件删除成功PATH=" + tmpFileName);
		}
		
		return bytes;
	}
	
	/**
	 * 按SHEETNAME解析工作薄
	 * 
	 * @param reportDatas 数据
	 * @param tmpFileName 临时模板文件
	 * @param sheetName
	 * @throws Exception
	 */
	protected void parseWorkBookBySheetName(ReportData reportData, String tmpFileName, String sheetName) throws Exception {
		
		this.parseWorkBookBySheetName(new Object[] { reportData.getBaseInfo(), reportData.getDatalist() }, tmpFileName, sheetName);
		
	}
	
	/**
	 * 按SHEETNAME解析工作薄
	 * 
	 * @param reportDatas 数据
	 * @param tmpFileName 临时模板文件
	 * @param sheetName
	 * @throws Exception
	 */
	protected void parseWorkBookBySheetName(Object[] reportDatas, String tmpFileName, String sheetName) throws Exception {
		
		logger.debug("打开EXCEL临时的模板文件PATH=" + tmpFileName);
		
		// 初始化EXCEL的SHEET
		excelTemplate.setCurrentSheetName(sheetName).executeTemplateParse(reportDatas);
		
		// 初始化工作薄取得相应的参数
		List<String> excelParams = excelTemplate.readExcelParams();
		
		ExcelConfiguration conf = new ExcelConfiguration();
		
		Map<String, Object> content = conf.fillData(excelParams, reportDatas);
		
		logger.debug("替换模板数据。");
		
		excelTemplate.replaceValue(content);
		
	}
	
	public void saveToFile(String tmpFileName) throws Exception {
		excelTemplate.saveToFile(tmpFileName);
	}
	
	/**
	 * Mapped File way MappedByteBuffer 可以在处理大文件时，提升性能
	 * 
	 * @param filename
	 * @return
	 * @throws IOException
	 */
	@SuppressWarnings("resource")
	public byte[] getBytes(String filename) throws IOException {
		
		FileChannel fc = null;
		
		try {
			fc = new RandomAccessFile(filename, "r").getChannel();
			MappedByteBuffer byteBuffer = fc.map(MapMode.READ_ONLY, 0, fc.size()).load();
			byte[] result = new byte[(int) fc.size()];
			if (byteBuffer.remaining() > 0) {
				byteBuffer.get(result, 0, byteBuffer.remaining());
			}
			return result;
		} catch (IOException e) {
			logger.debug(e.getMessage(), e);
			throw e;
		} finally {
			try {
				fc.close();
			} catch (IOException e) {
				logger.debug(e.getMessage(), e);
			}
		}
	}
	
	/**
	 * the traditional io way
	 * 
	 * @param filename
	 * @return
	 * @throws IOException
	 */
	public byte[] toByteArray(String filename) throws IOException {
		
		File f = new File(filename);
		if (!f.exists()) {
			throw new FileNotFoundException(filename);
		}
		
		ByteArrayOutputStream bos = new ByteArrayOutputStream((int) f.length());
		BufferedInputStream in = null;
		try {
			in = new BufferedInputStream(new FileInputStream(f));
			int buf_size = 1024;
			byte[] buffer = new byte[buf_size];
			int len = 0;
			while (-1 != (len = in.read(buffer, 0, buf_size))) {
				bos.write(buffer, 0, len);
			}
			return bos.toByteArray();
		} catch (IOException e) {
			logger.debug(e.getMessage(), e);
			throw e;
		} finally {
			try {
				in.close();
			} catch (IOException e) {
				logger.debug(e.getMessage(), e);
			}
			bos.close();
		}
	}
	
	/**
	 * 创建模板文件
	 * 
	 * @param templatePath 模板路径
	 * @param tempDictionaryPath 临时路径
	 * @return 复制的模板文件
	 * @throws Exception
	 */
	public String createTempTemplateFile(String templatePath, String tempDictionaryPath) throws Exception {
		
		String tmpFileName = "";
		
		// 生成下载文件的临时路径
		String tmpFilePath = tempDictionaryPath + File.separator + DateUtil.dateToString(new Date(), "yyyy-MM-dd") + File.separator;
		
		File file = new File(tmpFilePath);
		if (!file.exists()) {
			file.mkdirs();
		}
		
		// tmpFileName = tmpFilePath + UUID.randomUUID().toString() + ".xls";
		tmpFileName = tmpFilePath + UUID.randomUUID().toString() + ".et";
		
		// 复制一个文件去临时目录
		FileUtils.copyFile(new File(templatePath), new File(tmpFileName));
		
		excelTemplate.initializeWorkbook(tmpFileName);
		
		return tmpFileName;
		
	}
	
	/**
	 * 复制相应的SHEET
	 * 
	 * @param tmpFileName 带路径的模板的名称
	 * @param templateSheetName 模板SHEET名称
	 * @param sheetNames 要生成的SHEET名称
	 * @throws Exception
	 */
	public void cloneTemplateSheet(String tmpFileName, String templateSheetName, String sheetName, Integer position) throws Exception {
		
		logger.debug("开始克隆相应的SHEET。");
		
		excelTemplate.cloneSheet(templateSheetName, sheetName);
		
		excelTemplate.saveToFile(tmpFileName);
		
		excelTemplate = new ExcelTemplate();
		
		excelTemplate.initializeWorkbook(tmpFileName);
		
		logger.debug("克隆相应的SHEET结束。");
		
	}
	
	/**
	 * 复制相应的SHEET
	 * 
	 * @param tmpFileName 带路径的模板的名称
	 * @param templateSheetName 模板SHEET名称
	 * @param sheetNames 要生成的SHEET名称
	 * @throws Exception
	 */
	public void cloneTemplateSheet(String tmpFileName, String templateSheetName, String[] sheetNames) throws Exception {
		
		logger.debug("开始克隆相应的SHEET。");
		
		for (String sheetName : sheetNames) {
			excelTemplate.cloneSheet(templateSheetName, sheetName);
		}
		
		logger.debug("克隆相应的SHEET结束。");
		
	}
	
	/**
	 * 根据模板名称排序
	 * 
	 * @param excelTemplate
	 * @param sheetNames
	 */
	public void sortSheet(String[] sheetNames) {
		
		for (int i = 0; i < sheetNames.length; i++) {
			
			excelTemplate.sortSheet(sheetNames[i], i);
		}
	}
	
	/**
	 * 根据SHEET名后缀删除相应的SHEET
	 * 
	 * @param sheetNameSuffix
	 * @throws Exception
	 */
	public void removeSheetBySuffix(String tmpFileName, String sheetNameSuffix) throws Exception {
		
		excelTemplate.saveToFile(tmpFileName);
		excelTemplate = new ExcelTemplate();
		excelTemplate.initializeWorkbook(tmpFileName);
		excelTemplate.removeSheetBySuffix(sheetNameSuffix);
		excelTemplate.saveToFile(tmpFileName);
	}
	
}
