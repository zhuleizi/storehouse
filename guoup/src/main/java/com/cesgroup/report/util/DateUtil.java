package com.cesgroup.report.util;

import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DateUtil {
	public static final String FORMAT_TYPE_DAY = "yyyy-MM-dd";
	public static final String FORMAT_TYPE_MINUTE = "yyyy-MM-dd HH:mm";
	public static final String FORMAT_TYPE_Y_M_D_H_M_S = "yyyy-MM-dd HH:mm:ss";
	
	public static java.util.Date strToDate(String strDate)
			throws ParseException {
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		java.util.Date date = null;
		if (!strDate.equals("")) {
			date = format.parse(strDate);
		}
		return date;
	}
	
	public static java.util.Date strToDate(String strDate, String dateStyle)
			throws ParseException {
		SimpleDateFormat format = new SimpleDateFormat(dateStyle);
		java.util.Date date = null;
		if ((!"".equals(strDate)) && (strDate != null)) {
			date = format.parse(strDate);
		}
		return date;
	}
	
	public static String dateToString(java.util.Date date) {
		String strDate = "";
		if (date != null) {
			SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
			strDate = format.format(date);
		}
		return strDate;
	}
	
	public static String dateToString(java.util.Date date, String dateStyle)
			throws Exception {
		SimpleDateFormat format = null;
		String dateStr = "";
		try {
			format = new SimpleDateFormat(dateStyle);
			dateStr = format.format(date);
		} catch (Exception ex) {
			throw new Exception("Error occurs when format date to string type [" + dateStyle + "]：" + ex.getMessage());
		}
		
		return dateStr;
	}
	
	public static String dateToString(Timestamp date, String dateStyle)
			throws Exception {
		SimpleDateFormat format = null;
		String dateStr = "";
		try {
			format = new SimpleDateFormat(dateStyle);
			dateStr = format.format(date);
		} catch (Exception ex) {
			throw new Exception("Error occurs when format date to string type [" + dateStyle + "]：" + ex.getMessage());
		}
		
		return dateStr;
	}
	
	public static String dateToString(java.sql.Date date, String dateStyle)
			throws Exception {
		SimpleDateFormat format = null;
		String dateStr = "";
		try {
			format = new SimpleDateFormat(dateStyle);
			dateStr = format.format(date);
		} catch (Exception ex) {
			throw new Exception("Error occurs when format date to string type [" + dateStyle + "]：" + ex.getMessage());
		}
		
		return dateStr;
	}
	
	public static boolean isValidateDate(String date)
			throws Exception {
		boolean result = false;
		String eL = "^((\\d{2}(([02468][048])|([13579][26]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])))))|(\\d{2}(([02468][1235679])|([13579][01345789]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|(1[0-9])|(2[0-8]))))))";
		Pattern p = null;
		Matcher m = null;
		try {
			p = Pattern.compile(eL);
			m = p.matcher(date);
			result = m.matches();
		} catch (Exception ex) {
			throw new Exception("Error occurs when validate whether string is right date format：" + ex.getMessage());
		}
		
		return result;
	}
	
	public static java.util.Date getFirstDayOfNextNumberMonth(java.util.Date date, int NextNumber) {
		Calendar calendar1 = Calendar.getInstance();
		calendar1.setTime(date);
		calendar1.add(2, NextNumber);
		calendar1.set(5, 1);
		
		return calendar1.getTime();
	}
	
	public static java.util.Date getFirstDayOfCurrentMonth(java.util.Date date) {
		try {
			if (date == null) {
				return null;
			}
			Calendar calendar = Calendar.getInstance();
			calendar.setTime(date);
			calendar.set(5, 1);
			return calendar.getTime();
		} catch (Exception e) {
		}
		return null;
	}
	
	public static java.util.Date getFirstDayOfSysMonth(java.util.Date date) {
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(date);
		calendar.set(5, 1);
		calendar.set(11, 0);
		calendar.set(12, 0);
		calendar.set(13, 0);
		return calendar.getTime();
	}
	
	public static boolean isSameMonth(java.util.Date date1, java.util.Date date2) {
		boolean flg = false;
		int subYear = 0;
		Calendar calendar1 = Calendar.getInstance();
		calendar1.setTime(date1);
		
		Calendar calendar2 = Calendar.getInstance();
		calendar2.setTime(date1);
		
		subYear = calendar1.get(1) - calendar2.get(1);
		if ((subYear == 0) &&
				(calendar1.get(2) == calendar2.get(2))) {
			flg = true;
		}
		
		return flg;
	}
	
	public static java.util.Date getLastDayOfCurrentMonth(java.util.Date date) {
		try {
			if (date == null) {
				return null;
			}
			Calendar calendar = Calendar.getInstance();
			calendar.setTime(date);
			calendar.set(5, calendar.getActualMaximum(5));
			return calendar.getTime();
		} catch (Exception e) {
		}
		return null;
	}
	
	public static java.util.Date getLastDayOfNextNumberMonth(java.util.Date date, int nextNumberMonth) {
		try {
			if (date == null) {
				return null;
			}
			Calendar calendar = Calendar.getInstance();
			calendar.setTime(date);
			calendar.add(2, nextNumberMonth);
			calendar.set(5, calendar.getActualMaximum(5));
			return calendar.getTime();
		} catch (Exception e) {
		}
		return null;
	}
	
	public static String getNowYear() {
		java.util.Date currentTime = Calendar.getInstance().getTime();
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String dateString = formatter.format(currentTime);
		
		String year = dateString.substring(0, 4);
		return year;
	}
	
	public static String getNowHour() {
		java.util.Date currentTime = Calendar.getInstance().getTime();
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String dateString = formatter.format(currentTime);
		
		String hour = dateString.substring(11, 13);
		return hour;
	}
	
	public static String getNowMinute() {
		java.util.Date currentTime = Calendar.getInstance().getTime();
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String dateString = formatter.format(currentTime);
		
		String min = dateString.substring(14, 16);
		return min;
	}
	
	public static java.util.Date convertString2Date(Object strDate) {
		SimpleDateFormat sdf = null;
		java.util.Date date = null;
		String tmpValue = String.valueOf(strDate);
		if (null == tmpValue) {
			return date;
		}
		if (tmpValue.matches("[0-9]{2,4}-[0-9]{1,2}-[0-9]{1,2} [0-9]{1,2}:[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
		} else if (tmpValue.matches("[0-9]{2,4}-[0-9]{1,2}-[0-9]{1,2} [0-9]{1,2}:[0-9]{1,2}:[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		} else if (tmpValue.matches("[0-9]{2,4}-[0-9]{1,2}-[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("yyyy-MM-dd");
		} else if (tmpValue.matches("[0-9]{2,4}[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("yyyyMM");
		} else if (tmpValue.matches("[0-9]{2,4}[0-9]{1,2}[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("yyyyMMdd");
		} else if (tmpValue.matches("[0-9]{1,2}:[0-9]{1,2}:[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("HH:mm:ss");
		} else if (tmpValue.matches("[0-9]{2,4}[0-9]{1,2}[0-9]{1,2}[0-9]{1,2}[0-9]{1,2}[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("yyyyMMddHHmmss");
		} else if (tmpValue.matches("[0-9]{2,4}/[0-9]{1,2}/[0-9]{1,2}")) {
			sdf = new SimpleDateFormat("yyyy/MM/dd");
		}
		if (sdf == null)
			return null;
		try {
			date = sdf.parse(tmpValue);
		} catch (ParseException e) {
		}
		return date;
	}
	
	public static java.util.Date convertString2Date(String strDate, String dateStyle) {
		try {
			return strToDate(strDate, dateStyle);
		} catch (ParseException e) {
		}
		return null;
	}
	
	public static java.util.Date convertDate(Object date) {
		if ((date instanceof java.util.Date)) {
			return (java.util.Date) date;
		}
		return convertString2Date(date);
	}
	
	public static String getFirstDayOfMonth(String sDate, String eDate, String format) {
		try {
			if ((StringUtils.isEmpty(sDate)) && (StringUtils.isEmpty(eDate))) {
				java.util.Date date = getFirstDayOfCurrentMonth(new java.util.Date());
				String temp = new SimpleDateFormat("yyyy-MM-dd").format(date);
				if ("yyyy-MM-dd HH:mm".equals(format))
					temp = temp + " 00:00";
				return temp;
			}
		} catch (Exception e) {
		}
		return sDate;
	}
	
	public static int dateDiff(java.util.Date aDate, java.util.Date anotherDate) {
		int dates = (int) ((aDate.getTime() - anotherDate.getTime()) / 86400000L);
		return dates;
	}
	
	public static java.util.Date addDay(java.util.Date date, int day) {
		if (null == date) {
			return null;
		}
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(date);
		calendar.set(6, calendar.get(6) + day);
		return calendar.getTime();
	}
	
	public static java.util.Date addMinute(java.util.Date date, int minute) {
		GregorianCalendar calendar = new GregorianCalendar();
		calendar.setTime(date);
		calendar.add(12, minute);
		return calendar.getTime();
	}
	
	public static java.util.Date addSecond(java.util.Date date, int second) {
		GregorianCalendar calendar = new GregorianCalendar();
		calendar.setTime(date);
		calendar.add(13, second);
		return calendar.getTime();
	}
	
	public static java.util.Date subDay(java.util.Date date, int day) {
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(date);
		calendar.set(6, calendar.get(6) - day);
		return calendar.getTime();
	}
	
	public static String formatDate2String(java.util.Date date, String pattern) {
		SimpleDateFormat df = null;
		String returnValue = "";
		
		if (date != null) {
			df = new SimpleDateFormat(pattern);
			returnValue = df.format(date);
		}
		return returnValue;
	}
	
	public static boolean isWorkDay(java.util.Date date) {
		Calendar calendar = Calendar.getInstance();
		int week = calendar.get(7) - 1;
		
		return (week > 0) && (week < 6);
	}
	
	public static boolean isToday(java.util.Date date) {
		String today = formatDate2String(new java.util.Date(), "yyyy-MM-dd");
		
		String compareDate = formatDate2String(date, "yyyy-MM-dd");
		
		return StringUtils.equals(today, compareDate);
	}
	
	public static long difference(long day1, long day2) {
		long diff = day2 - day1;
		return diff / 1000L / 60L;
	}
	
	public static List<java.util.Date> getCurrentMonthAllDay(java.util.Date firstCurrentMonthDay) {
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(firstCurrentMonthDay);
		
		int currentMonth = calendar.get(2);
		List days = new ArrayList();
		for (int i = 1; (i < 32) &&
				(currentMonth == calendar.get(2)); i++) {
			days.add(calendar.getTime());
			calendar.set(6, calendar.get(6) + 1);
		}
		return days;
	}
	
	public static String getYearMonth(Object strDate, int month) {
		java.util.Date date = convertDate(strDate);
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(date);
		calendar.add(2, month);
		return calendar.get(1) * 100 + calendar.get(2) + 1 + "";
	}
	
	public static int getDayOfMonth() {
		Calendar calendar = Calendar.getInstance();
		return calendar.get(5);
	}
	
	public static java.util.Date padTime(java.util.Date date) {
		if (null != date) {
			String strTime = formatDate2String(date, "HH:mm:ss");
			if (!StringUtils.eq("00:00:00", strTime)) {
				return date;
			}
			String strDate = formatDate2String(date, "yyyy-MM-dd");
			String strSystemDate = formatDate2String(new java.util.Date(), "yyyy-MM-dd");
			if (StringUtils.eq(strDate, strSystemDate)) {
				String strSystemTime = formatDate2String(new java.util.Date(), "HH:mm:ss");
				return convertDate(strDate + " " + strSystemTime);
			}
		}
		return date;
	}
	
	public static String formatDate(java.util.Date date, String format) {
		try {
			return dateToString(date, format);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "";
	}
	
	public static List<String> getMonthBetween(java.util.Date startDate, java.util.Date endDate, String dateStyle)
			throws ParseException {
		ArrayList result = new ArrayList();
		SimpleDateFormat sdf = new SimpleDateFormat(dateStyle);
		
		Calendar min = Calendar.getInstance();
		Calendar max = Calendar.getInstance();
		
		min.setTime(startDate);
		min.set(min.get(1), min.get(2), 1);
		
		max.setTime(endDate);
		max.set(max.get(1), max.get(2), 2);
		
		Calendar curr = min;
		while (curr.before(max)) {
			result.add(sdf.format(curr.getTime()));
			curr.add(2, 1);
		}
		
		return result;
	}
}