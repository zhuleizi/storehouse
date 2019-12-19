package com.cesgroup.report.util;

import java.io.UnsupportedEncodingException;
import java.math.BigDecimal;
import java.security.InvalidParameterException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.StringTokenizer;
import java.util.regex.Pattern;

public class StringUtils {
	public static final StringFormatConstants LOWER_CASE = StringFormatConstants.LOWER_CASE;
	
	public static final StringFormatConstants UPPER_CASE = StringFormatConstants.UPPER_CASE;
	private static final String EMPTY_STRING = "";
	public static final String HALF_SPACE_STRING = " ";
	private static final String WIDECHARS = "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ１２３４５６７８９０";
	private static final String CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890";
	private static final char FULL_SPACE = '　';
	private static final char HALF_SPACE = ' ';
	private static final String PUNC_BRACKET = "[";
	private static final String ESCAPE_BRACKET = "[[]";
	private static final String PUNC_UNDERLINE = "_";
	private static final String ESCAPE_UNDERLINE = "[_]";
	private static final String PUNC_PERCENT = "%";
	private static final String ESCAPE_PERCENT = "[%]";
	private static final String PUNC_SINGLE_QUOTE = "'";
	private static final String ESCAPE_SINGLE_QUOTE = "''";
	private static final String NUMBER_REGEX = "[0-9]*";
	private static final String MAIL_REGEX = "\\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
	public static final String HHMM_REGEX = "^(0\\d{1}|1\\d{1}|2[0-3]):([0-5]\\d{1})$";
	private static final String SPACE_REGEX = ".*\\s.*";
	private static StringFormatType[] dateParsePattern = { StringFormatType.DATE_FORMAT_DATETIME, StringFormatType.DATE_FORMAT_TIME, StringFormatType.DATE_FORMAT_YEAR_MON_DAY,
			StringFormatType.DATE_FORMAT_YEAR_MON, StringFormatType.DATE_FORMAT_FYEAR_MON, StringFormatType.DATE_FORMAT_YY_MM_DD, StringFormatType.DATE_FORMAT_YY_M_D,
			StringFormatType.DATE_FORMAT_YYYY_MM_DD, StringFormatType.DATE_FORMAT_YYYY_M_D, StringFormatType.DATE_FORMAT_YY__MM__DD, StringFormatType.DATE_FORMAT_YY__M__D,
			StringFormatType.DATE_FORMAT_YYYY__MM__DD, StringFormatType.DATE_FORMAT_YYYY__M__D, StringFormatType.DATE_FORMAT_YYYY__MM__DD__HH_MM_SS_SSS, StringFormatType.DATE_FORMAT_YYYYMMDD,
			StringFormatType.DATE_FORMAT_YYYYMM, StringFormatType.DATE_FORMAT_YYYYMMDDHHMMSS, StringFormatType.DATE_FORMAT_YYYY_MM, StringFormatType.DATE_FORMAT_YYMM };
	
	private static StringFormatType[] numberParsePattern = { StringFormatType.NUMBER_FORMAT_INTEGER_NUMBER, StringFormatType.NUMBER_FORMAT_MONEY, StringFormatType.NUMBER_FORMAT_NUMBER,
			StringFormatType.NUMBER_FORMAT_NORMAL_NUMBER, StringFormatType.NUMBER_FORMAT_PER_NUMBER, StringFormatType.NUMBER_FORMAT_00_NUMBER };
	
	public static String append(String str, char c, int len) {
		return rappend(lappend(str, c, len), c, len);
	}
	
	public static String camelToUnderline(String obj) {
		assertNotNull(obj);
		
		String[] subStrings = obj.split("[A-Z]");
		
		String retval = "";
		
		for (int i = 0; i < subStrings.length - 1; i++) {
			int idx = obj.indexOf(subStrings[i], retval.length() - i) + subStrings[i].length();
			
			retval = retval + subStrings[i].toUpperCase() + "_" + obj.substring(idx, idx + 1);
		}
		
		retval = retval + subStrings[(subStrings.length - 1)].toUpperCase();
		
		if ((0 == retval.indexOf("_")) && (1 < retval.length())) {
			retval = retval.substring(1);
		}
		
		return retval;
	}
	
	public static String firstLetterFormat(String obj, StringFormatConstants type) {
		assertNotNull(type);
		
		String firstLetter = null;
		
		if (isEmpty(obj)) {
			return obj;
		}
		
		if (StringFormatConstants.UPPER_CASE.equals(type)) {
			firstLetter = obj.substring(0, 1).toUpperCase();
		} else if (StringFormatConstants.LOWER_CASE.equals(type)) {
			firstLetter = obj.substring(0, 1).toLowerCase();
		}
		
		if (1 < obj.length()) {
			return firstLetter + obj.substring(1);
		}
		return firstLetter;
	}
	
	public static String formatDate(Date dt, StringFormatType formatType) {
		assertNotNull(formatType);
		
		if (null == dt) {
			return null;
		}
		
		SimpleDateFormat formatter = new SimpleDateFormat(formatType.format());
		return formatter.format(dt);
	}
	
	public static String formatNumber(BigDecimal bd, StringFormatType formatType) {
		assertNotNull(formatType);
		
		if (null == bd) {
			return null;
		}
		
		DecimalFormat formatter = new DecimalFormat(formatType.format());
		return formatter.format(bd);
	}
	
	public static boolean isEmpty(String obj) {
		if (null == obj) {
			return true;
		}
		return "".equals(obj);
	}
	
	public static String lalign(String str, char c, int len) {
		assertNotNull(str);
		
		if (str.length() >= len) {
			return str;
		}
		
		return str + repeat(c, len - str.length());
	}
	
	public static String lalign(String str, int len) {
		return lalign(str, ' ', len);
	}
	
	public static String lappend(String str, char c, int len) {
		assertNotNull(str);
		assertNotNegative(len);
		
		return repeat(c, len) + str;
	}
	
	public static String ltrim(String str) {
		assertNotNull(str);
		
		for (int i = 0; i < str.length(); i++) {
			if ((' ' != str.charAt(i)) && ('　' != str.charAt(i))) {
				return str.substring(i);
			}
		}
		return "";
	}
	
	public static String ltrim(String str, char c) {
		assertNotNull(str);
		
		for (int i = 0; i < str.length(); i++) {
			if (str.charAt(i) != c) {
				return str.substring(i);
			}
		}
		return "";
	}
	
	public static Date parseDate(String dateString) {
		assertNotNull(dateString);
		
		SimpleDateFormat parser = new SimpleDateFormat();
		
		Date retval = null;
		
		boolean chkResult = false;
		
		for (StringFormatType pattern : dateParsePattern) {
			try {
				parser.applyPattern(pattern.format());
				retval = parser.parse(dateString);
				
				if (dateString.equals(formatDate(retval, pattern))) {
					chkResult = true;
					break;
				}
			} catch (ParseException pe) {
			}
		}
		if (!chkResult) {
			retval = null;
		}
		
		return retval;
	}
	
	public static BigDecimal parseNumber(String numberString) {
		assertNotNull(numberString);
		
		DecimalFormat parser = new DecimalFormat();
		
		parser.setParseBigDecimal(true);
		
		BigDecimal retval = null;
		
		boolean chkResult = false;
		
		for (StringFormatType pattern : numberParsePattern) {
			try {
				parser.applyPattern(pattern.format());
				retval = (BigDecimal) parser.parse(numberString);
				if (numberString.equals(formatNumber(retval, pattern))) {
					chkResult = true;
					break;
				}
			} catch (ParseException pe) {
			}
		}
		if (!chkResult) {
			retval = null;
		}
		
		return retval;
	}
	
	public static String ralign(String str, char c, int len) {
		assertNotNull(str);
		assertNotNegative(len);
		
		if (str.length() >= len) {
			return str;
		}
		
		return repeat(c, len - str.length()) + str;
	}
	
	public static String ralign(String str, int len) {
		return ralign(str, ' ', len);
	}
	
	public static String rappend(String str, char c, int len) {
		assertNotNull(str);
		assertNotNegative(len);
		
		return str + repeat(c, len);
	}
	
	public static String repeat(char c, int len) {
		assertNotNegative(len);
		
		char[] carr = new char[len];
		for (int i = 0; i < len; i++) {
			carr[i] = c;
		}
		return new String(carr);
	}
	
	public static String rtrim(String str) {
		assertNotNull(str);
		
		for (int i = str.length() - 1; i >= 0; i--) {
			if ((' ' != str.charAt(i)) && ('　' != str.charAt(i))) {
				return str.substring(0, i + 1);
			}
		}
		return "";
	}
	
	public static String rtrim(String str, char c) {
		assertNotNull(str);
		
		for (int i = str.length() - 1; i >= 0; i--) {
			if (str.charAt(i) != c) {
				return str.substring(0, i + 1);
			}
		}
		return "";
	}
	
	public static String sqlEscape(String str) {
		assertNotNull(str);
		
		return str.replace("[", "[[]").replace("%", "[%]").replace("_", "[_]").replace("'", "''");
	}
	
	public static String sqlUnescape(String str) {
		assertNotNull(str);
		
		return str.replace("''", "'").replace("[_]", "_").replace("[%]", "%").replace("[[]", "[");
	}
	
	public static String trim(String str) {
		return ltrim(rtrim(str));
	}
	
	public static String trim(String str, char c) {
		return rtrim(ltrim(str, c), c);
	}
	
	public static String underlineToCamel(String obj) {
		assertNotNull(obj);
		
		String[] subStrings = obj.split("_");
		
		String retval = "";
		
		for (String subString : subStrings) {
			retval = retval + firstLetterFormat(subString.toLowerCase(), StringFormatConstants.UPPER_CASE);
		}
		
		return retval;
	}
	
	public static String replaceWideChars(String input) {
		StringBuffer output = new StringBuffer(input.length());
		for (int i = 0; i < input.length(); i++) {
			int idx = "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ１２３４５６７８９０".indexOf(input.charAt(i));
			if (idx != -1)
				output.append("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890".charAt(idx));
			else {
				output.append(input.charAt(i));
			}
		}
		return output.toString();
	}
	
	public static String replaceHalfChars(String input) {
		StringBuffer output = new StringBuffer(input.length());
		for (int i = 0; i < input.length(); i++) {
			int idx = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890".indexOf(input.charAt(i));
			if (idx != -1)
				output.append("ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ１２３４５６７８９０".charAt(idx));
			else {
				output.append(input.charAt(i));
			}
		}
		return output.toString();
	}
	
	public static String removePercent(String input) {
		String strObj = null;
		if ((!isEmpty(input)) && (input.startsWith("%")) && (input.endsWith("%"))) {
			strObj = input.substring(1);
			strObj = strObj.substring(0, strObj.length() - 1);
		}
		return strObj;
	}
	
	public static String nullToSp(String str) {
		if (isEmpty(str)) {
			return "";
		}
		return str;
	}
	
	public static boolean isNumeirc(String str) {
		Pattern pattern = Pattern.compile("[0-9]*");
		return pattern.matcher(str).matches();
	}
	
	public static boolean isFormatNormal(String target, String regex) {
		Pattern pattern = Pattern.compile(regex);
		return pattern.matcher(target).matches();
	}
	
	public static boolean isSimpleMailFormat(String str) {
		Pattern pattern = Pattern.compile("\\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*");
		return pattern.matcher(str).matches();
	}
	
	public static boolean isContainSpace(String str) {
		Pattern pattern = Pattern.compile(".*\\s.*");
		return pattern.matcher(str).matches();
	}
	
	private static void assertNotNull(Object obj) {
		if (null == obj)
			throw new InvalidParameterException();
	}
	
	private static void assertNotNegative(int num) {
		if (0 > num)
			throw new InvalidParameterException();
	}
	
	public static String clearSameData(String data) {
		if (isNotEmpty(data)) {
			List tmp = new ArrayList();
			String[] arr = data.split(",");
			for (String str : arr) {
				if (tmp.contains(str))
					continue;
				tmp.add(str);
			}
			
			if (!tmp.isEmpty()) {
				return tmp.toString().substring(1, tmp.toString().length() - 1).replaceAll(" ", "");
			}
		}
		
		return "";
	}
	
	public static boolean isNotEmpty(String str) {
		return !isEmpty(str);
	}
	
	public static boolean equals(String str, String str1) {
		if ((isNotEmpty(str)) && (isNotEmpty(str1))) {
			return str.equals(str1);
		}
		return false;
	}
	
	public static boolean equalsArray(String str, String[] arr) {
		for (String string : arr) {
			if ((isNotEmpty(str)) && (isNotEmpty(string)) &&
					(str.equals(string))) {
				return true;
			}
			
		}
		
		return false;
	}
	
	public static boolean equalsEmptyOrNull(String str, String str1) {
		boolean flag = equals(str, str1);
		if ((!flag) &&
				(isEmpty(str)) && (isEmpty(str1))) {
			flag = true;
		}
		
		return flag;
	}
	
	public static String format(String msg, String[] params) {
		if ((null != params) && (isNotEmpty(msg))) {
			for (int i = 0; i < params.length; i++) {
				String replacement = isNotEmpty(params[i]) ? params[i] : "";
				msg = msg.replaceAll("\\{" + i + "\\}", replacement);
			}
		}
		return msg;
	}
	
	public static String byteToString(byte[] content) {
		if ((content != null) && (content.length > 0)) {
			return new String(content);
		}
		return null;
	}
	
	public static String changeCharset(String str, String newCharset) {
		if (str != null) {
			byte[] bs = str.getBytes();
			try {
				return new String(bs, newCharset);
			} catch (UnsupportedEncodingException e) {
				e.printStackTrace();
			}
		}
		return null;
	}
	
	public static byte[] stringToByte(String content) {
		if ((content != null) && (!"".equals(content))) {
			return content.getBytes();
		}
		return null;
	}
	
	public static String arrConvertString(Object[] arr, String symbol) {
		StringBuffer buffer = new StringBuffer();
		if ((null != arr) && (arr.length > 0)) {
			for (Object obj : arr) {
				if (buffer.length() > 0) {
					buffer.append(symbol);
				}
				buffer.append(obj);
			}
		}
		return buffer.toString();
	}
	
	public static String firstToUpper(String str) {
		return str.substring(0, 1).toUpperCase() + str.substring(1);
	}
	
	public static String[] split(String str, String delim) {
		StringTokenizer token = new StringTokenizer(str, delim);
		String[] params = new String[token.countTokens()];
		for (int i = 0; i < params.length; i++) {
			params[i] = token.nextToken();
		}
		return params;
	}
	
	public static boolean eq(String str, String str1) {
		if ((isNotEmpty(str)) && (isNotEmpty(str1))) {
			return str.equals(str1);
		}
		return (isEmpty(str)) && (isEmpty(str1));
	}
	
	public static String delTrim(String str) {
		String strs = "";
		if ((str != null) && (str.length() > 0)) {
			return str.replaceAll(" ", "");
		}
		return strs;
	}
	
	public static String delLinefeed(String str) {
		String strs = "";
		if ((str != null) && (str.length() > 0)) {
			return str.replaceAll("\r\n", " ");
		}
		return strs;
	}
	
	public static boolean isExcel(String excelType) {
		return ("application/msexcel".equals(excelType)) || ("application/vnd.ms-excel".equals(excelType)) || ("application/octet-stream".equals(excelType)) || ("application/kset".equals(excelType))
				|| ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".equals(excelType));
	}
	
	public static enum StringFormatType {
		NULL("#,##0.00"),
		
		DATE_FORMAT_DATETIME("yyyy/MM/dd HH:mm:ss"),
		
		DATE_FORMAT_TIME("HH:mm:ss"),
		
		DATE_FORMAT_YEAR_MON_DAY("yyyy年MM月dd日"),
		
		DATE_FORMAT_YEAR_MON_DAY_SINGLE("yyyy年M月d日"),
		
		DATE_FORMAT_MON_DAY("MM月dd日"),
		
		DATE_FORMAT_YEAR_MON("yy年MM月"),
		
		DATE_FORMAT_FYEAR_MON("yyyy年MM月"),
		
		DATE_FORMAT_TYEAR_MON("yy年MM月"),
		
		DATE_FORMAT_YY_MM_DD("yy/MM/dd"),
		
		DATE_FORMAT_YY_M_D("yy/M/d"),
		
		DATE_FORMAT_YYYY_MM_DD("yyyy/MM/dd"),
		
		DATE_FORMAT_YYYY_M_D("yyyy/M/d"),
		
		DATE_FORMAT_YY__MM__DD("yy-MM-dd"),
		
		DATE_FORMAT_YY__M__D("yy-M-d"),
		
		DATE_FORMAT_YYYY__MM__DD("yyyy-MM-dd"),
		
		DATE_FORMAT_YYYY__M__D("yyyy-M-d"),
		
		DATE_FORMAT_YYYY__MM__DD__HH_MM("yyyy/MM/dd HH:mm"),
		
		DATE_FORMAT_YYYY__MM__DD__HH_MM_SS("yyyy/MM/dd HH:mm:ss"),
		
		DATE_FORMAT_YYYY__MM__DD__HH_MM_SS_SSS("yyyy-MM-dd HH:mm:ss.SSS"),
		
		DATE_FORMAT_YYYYMMDD("yyyyMMdd"),
		
		DATE_FORMAT_YYYYMM("yyyyMM"),
		
		DATE_FORMAT_YYYYMMDDHHMMSS("yyyyMMddHHmmss"),
		
		DATE_FORMAT_YYYYMMDDHHMMSSS("yyyyMMddHHmmsss"),
		
		DATE_FORMAT_YYYYMMDDHHMMSSSSS("yyyyMMddHHmmssSSS"),
		
		DATE_FORMAT_YYYY_MM("yyyy/MM"),
		
		DATE_FORMAT_HH_MM("yyyy/MM/dd HH:mm"),
		
		DATE_FORMAT_MM_DD("MM/dd"),
		
		DATE_FORMAT_YYMM("yyMM"),
		
		DATE_FORMAT_MMDD("MMdd"),
		
		NUMBER_FORMAT_INTEGER_NUMBER("#,##0"),
		
		NUMBER_FORMAT_MONEY("#,##0.00"),
		
		NUMBER_FORMAT_MONEY_JPY("#,##0"),
		
		NUMBER_FORMAT_FOURTEEN_MONEY("#,##0.00############"),
		
		NUMBER_FORMAT_MONEY_NUMBER("#,##0.0"),
		
		NUMBER_FORMAT_NUMBER("0.00"),
		
		NUMBER_FORMAT_PER_NUMBER("0.000"),
		
		NUMBER_FORMAT_FOUR_NUMBER("0.0000"),
		
		NUMBER_FORMAT_FIVE_NUMBER("0.00000"),
		
		NUMBER_FORMAT_SIX_NUMBER("0.000000"),
		
		NUMBER_FORMAT_SEVEN_NUMBER("0.0000000"),
		
		NUMBER_FORMAT_EIGHT_NUMBER("0.00000000"),
		
		NUMBER_FORMAT_FOURTEEN_NUMBER("0.00############"),
		
		NUMBER_FORMAT_FIFTEEN_NUMBER("0.00#############"),
		
		NUMBER_FORMAT_NORMAL_NUMBER("0"),
		
		NUMBER_FORMAT_00_NUMBER("00");
		
		private String formatString;
		
		private StringFormatType(String formatString) {
			this.formatString = formatString;
		}
		
		public String format() {
			return this.formatString;
		}
	}
	
	private static enum StringFormatConstants {
		LOWER_CASE,
		
		UPPER_CASE;
	}
}