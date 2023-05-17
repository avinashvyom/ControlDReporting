package com.vyomlabs.exceltotextconversion;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

public class ExcelRead {
	public static void main(String[] args) {
		String filePath = args[0];
		System.out.println("File path : " + filePath);
		HashMap<String, List<JobDetails>> dataMap = new HashMap<>();
		try {
			FileInputStream fis = new FileInputStream(filePath);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			List<JobDetails> rowData = new ArrayList<>();
			List<String> distIDList = new ArrayList<>();
			ArrayList<String> tempJobData = new ArrayList<>();
			String distID = "";
			Row row;
			//PoiItemReader<T>
			//BeanWrapperRowMapper<T>
			Iterator<Row> rows = sheet.rowIterator();
			//System.out.println("Total no of rows : " + sheet.getLastRowNum());
			while (rows.hasNext()) {
				row = (Row) rows.next();
				for (int i = 0; i < row.getLastCellNum(); i++) {
					@SuppressWarnings("unused")
					Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
				}
				String key = null;
				List<JobDetails> jobData = new ArrayList<>();
				if (row.getRowNum() == 0) {
					continue;
				}
				tempJobData = new ArrayList<>();
				for (Cell cell : row) {
					distIDList = new ArrayList<>();
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);
					//System.out.println("Cell value : " + cell.getStringCellValue());
					//System.out.println("Cell address : " + cell.getRowIndex() + cell.getColumnIndex());
					key = row.getCell(0).getStringCellValue().trim();
					//System.out.println("Key is :" + key);
					if (XSSFCell.CELL_TYPE_STRING == cell.getCellType()) {
						if (cell.getColumnIndex() == 6) {
							distID = cell.getStringCellValue();
							//System.out.println("previous distid"+distID);
							//distID = distID1.trim();
							distID = distID.replaceAll("\u00A0", "");
							//System.out.println("final distID"+distID);
						} else {
							String item = cell.getStringCellValue();
							item = item.replaceAll("\u00A0", "");
							tempJobData.add(item);
						}
					}
				}
				if (distID.isEmpty()) {
					distIDList = new ArrayList<>();
				} else {
					if (distID.contains(",")) {
						distIDList = seprateDistIDs(distID);
					} else {
						distIDList.add(distID);
					}
				}
				JobDetails jobDetails = new JobDetails();
				jobDetails.setReportID(tempJobData.get(1).toString().trim());
				jobDetails.setStepName(tempJobData.get(2).toString().trim());
				jobDetails.setProcStep(tempJobData.get(3).toString().trim());
				jobDetails.setDDName(tempJobData.get(4).toString().trim());
				jobDetails.setClassName(tempJobData.get(5).toString().trim());
				jobDetails.setCategory(tempJobData.get(6));
				//jobDetails.setScheduleDescription(tempJobData.get(7));
				//jobDetails.setForm(tempJobData.get(6).toString());
				jobDetails.setDistID(distIDList);
				jobData.add(jobDetails);
				if (dataMap.containsKey(key)) {
					rowData = dataMap.get(key);
					rowData.addAll(jobData);
				} else {
					dataMap.put(key, jobData);
				}
			}
			fis.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
		FileWriter fw;
		BufferedWriter bw;
		
		for (Map.Entry<String, List<JobDetails>> entry : dataMap.entrySet()) {
			String key = entry.getKey();
			File file = new File(Paths.get("").toAbsolutePath().toString() + "/" + key);
			try {
				if (!file.exists()) {
					file.createNewFile();
					System.out.println("File created with name :" + file.getName() + ",at path :" + file.getAbsolutePath());
					fw = new FileWriter(file.getAbsoluteFile());
					bw = new BufferedWriter(fw);
					List<JobDetails> jobs = dataMap.get(entry.getKey());
					ArrayList<String> classList = new ArrayList<String>();
					ArrayList<String> scheduleList = new ArrayList<String>();
					HashMap<String, Integer> classIndexMap = new HashMap<String, Integer>();
					//Set<String> classSet = new HashSet<String>();
					//ArrayList<String> tempClassList = new ArrayList<String>();
					//ArrayList<String> tempScheduleList = new ArrayList<String>();			
					int index = 0;
					// set = {D,C}
					Collections.sort(jobs,(o1,o2) -> o1.getClassName().compareTo(o2.getClassName()));
					for(JobDetails job : jobs) {
						classIndexMap.put(job.getClassName(), index);
						index++;
						//classSet.add(job.getClassName());
					}
					index = 0;
					//System.out.println("Class set for job " + key + " is :" + classSet);
					
					for(JobDetails job : jobs) {
						if(classList.contains(job.getClassName()) && scheduleList.contains(job.getScheduleDescription())) {
							bw.write(writeDSNBlock(job));
						}
						else {
							classList.add(job.getClassName());
							scheduleList.add(job.getScheduleDescription());
							//System.out.println("Class List : " + classList);
							//System.out.println("Schedule List : " + scheduleList);
							//System.out.println("Writing Default Block..........");
							bw.write(writeDefaultBlock(key, job.getScheduleDescription(), job.getCategory()));
							//System.out.println("Writing Class Block..........");
							bw.write(writeClassBlock(job.getClassName(), entry.getKey()));
							//System.out.println("Writing DSN Block..........");
							bw.write(writeDSNBlock(job));

						}
						if(classIndexMap.get(job.getClassName()) == index) {
							bw.write(getSpaces(15) + "\n");
							bw.write("0I833568 2023032107.16.00I833568 2023032110.00.529.0.20    01.04");
						}
						index++;
						bw.write(getSpaces(15) + "\n");
					}
					bw.close();
					fw.close();
					RandomAccessFile file1 = new RandomAccessFile(file, "rw");
					long length = file1.length();
					if (length == 0) {
						file1.close();
						return;
					}
					long position = length - 1;
					file1.seek(position);
					int lastChar = file1.read();
					if (lastChar == 10 || lastChar == 13) {
						file1.setLength(position);
					}
					file1.close();
				}
				
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	private static List<String> seprateDistIDs(String distID) {
		String[] arr = distID.split(",");
		return Arrays.asList(arr);
	}

	private static char[] writeDefaultBlock(String key, String scheduleDescription, String category) {
		StringBuffer sb = new StringBuffer();
		sb.append("R");
		if(key.length()<8) {
			sb.append(key).append(getSpaces(8-(key.length())));
		}
		else {
			sb.append(key);
		}
		sb.append(getSpaces(20)).append("I833568").append(getSpaces(1)).append(category).append(getSpaces(20-category.length())).append("REP00").append(getSpaces(17)).append("\n");
		//sb.append(getSchedule(scheduleDescription));
		// if schedule wanted, uncomment the above line else append code on below 2 lines. 
		sb.append("M*M").append(getSpaces(8)).append("YYYYYYYYYYYY").append(getDefaultSchedule(31)).append(getSpaces(16)).append("305").append(getSpaces(6)).append("\n");
		sb.append("V").append(getSpaces(1)).append("W").append(getSpaces(8)).append("YYYYYYYYYYYY").append(getDefaultSchedule(7)).append(getSpaces(48)).append("A").append("\n");
				sb.append("Q").append(getSpaces(2)).append("99999999").append(getSpaces(68)).append("\n")
				.append("F").append(getSpaces(6)).append("UNIDENTA").append(getSpaces(64)).append("\n");
				
		//System.out.println(sb.toString());
		return sb.toString().toCharArray();		
	}

	@SuppressWarnings("unused")
	private static char[] getSchedule(String scheduleDescription) {
		// TODO Auto-generated method stub
		//StringBuffer sb = new StringBuffer();
		List<String> schedules = new ArrayList<String>();
		if(scheduleDescription.equals("")) {
			//System.out.println("No schedule");
			return getNulCharacter(scheduleDescription);
		}
		else if(scheduleDescription.equals("WORKDAY LAST")) {
			//System.out.println("WORKDAY LAST SCHEDULE");
			LastWorkingDay lastWorkingDay = new LastWorkingDay();
			schedules = lastWorkingDay.getLastWorkingDay(lastWorkingDay.getCurrentMonth());
			//System.out.println("Schedule from Last working day block : "+schedules);
			//List<Integer> monthList = new ArrayList<>();
			return prepareLastWorkingDaySchedule(schedules);
		}
		else {
			int day = 0;
			schedules = Arrays.asList(scheduleDescription.split(","));
			try {
				day = Integer.parseInt(schedules.get(0));
				List<Integer> monthList = new ArrayList<>();
				for(String schedule : schedules) {
					monthList.add(Integer.parseInt(schedule));
				}
				//System.out.println("Schedules are monthly...." + day);
				return getMonthSchedule(monthList);
			}
			catch(NumberFormatException e) {
				List<String> weekList = new ArrayList<>();
				weekList.addAll(schedules);
				System.out.println("Schedules are weekly....");
				return getWeekSchedule(weekList);
			}			
		}
	}

	private static char[] prepareLastWorkingDaySchedule(List<String> schedules) {
		HashMap<Integer, Character> tempSchedule = new HashMap<>();
		tempSchedule.put(0, ' ');
		tempSchedule.put(1, ' ');
		tempSchedule.put(2, ' ');
		tempSchedule.put(3, ' ');
		tempSchedule.put(4, ' ');
		tempSchedule.put(5, ' ');
		tempSchedule.put(6, ' ');
		// System.out.println("prepareLastWorkingDaySchedule day :
		// "+schedules.get(0) + "day of week : "+schedules.get(1));
		switch (schedules.get(1)) {
		case "SUNDAY":
			tempSchedule.put(0, '{');
			break;
		case "MONDAY":
			tempSchedule.put(1, '{');
			break;
		case "TUESDAY":
			tempSchedule.put(2, '{');
			break;
		case "WEDNESDAY":
			tempSchedule.put(3, '{');
			break;
		case "THURSDAY":
			tempSchedule.put(4, '{');
			break;
		case "FRIDAY":
			tempSchedule.put(5, '{');
			break;
		case "SATURDAY":
			tempSchedule.put(6, '{');
			break;
		}
		char[] lastWorkingDayWeekSchedule = new char[7];
		StringBuffer weekSchedule = new StringBuffer();
		for (int j = 0; j < lastWorkingDayWeekSchedule.length; j++) {
			lastWorkingDayWeekSchedule[j] = tempSchedule.get(j);
			weekSchedule.append(lastWorkingDayWeekSchedule[j]);
		}
		char[] lastWorkingDayMonthSchedule = new char[31];
		Arrays.fill(lastWorkingDayMonthSchedule, ' ');
		int lastWorkingDay = Integer.parseInt(schedules.get(0));
		lastWorkingDayMonthSchedule[lastWorkingDay - 1] = '{';
		StringBuffer monthSchedule = new StringBuffer();
		for (int k = 0; k < lastWorkingDayMonthSchedule.length; k++) {
			monthSchedule.append(lastWorkingDayMonthSchedule[k]);
		}
		StringBuffer finalLastWorkingDaySchedule = new StringBuffer();
		finalLastWorkingDaySchedule.append("M*M").append(getSpaces(8)).append("YYYYYYYYYYYY").append(monthSchedule.toString()).append(getSpaces(16)).append("305").append(getSpaces(6)).append("\n");
		finalLastWorkingDaySchedule.append("V").append(getSpaces(1)).append("W").append(getSpaces(8)).append("YYYYYYYYYYYY").append(weekSchedule.toString()).append(getSpaces(47)).append("YA").append("\n");

		return finalLastWorkingDaySchedule.toString().toCharArray();
	}

	private static char[] getWeekSchedule(List<String> schedules) {
		//LastWorkingDay lastWorkingDay = new LastWorkingDay();
		
		HashMap<Integer, String> tempSchedule = new HashMap<>();
		tempSchedule.put(0, "M");
		tempSchedule.put(1, "T");
		tempSchedule.put(2, "W");
		tempSchedule.put(3, "TH");
		tempSchedule.put(4, "F");
		tempSchedule.put(5, "S");
		tempSchedule.put(6, "SU");
		for (int key : tempSchedule.keySet()) {
			if (schedules.contains(tempSchedule.get(key).toString())) {
				tempSchedule.put(key, "{");
			} else {
				tempSchedule.put(key, " ");
			}
		}
		String[] finalSchedule = new String[7];
		StringBuffer weekSchedule = new StringBuffer();
		for (int j = 0; j < finalSchedule.length; j++) {
			finalSchedule[j] = tempSchedule.get(j);
			weekSchedule.append(finalSchedule[j]);
		}
		StringBuffer finalWeekSchedule = new StringBuffer();
		finalWeekSchedule.append("M*M").append(getSpaces(8)).append("YYYYYYYYYYYY").append(getDefaultSchedule(31, "")).append(getSpaces(16)).append("305").append(getSpaces(6)).append("\n");
		finalWeekSchedule.append("V").append(getSpaces(1)).append("W").append(getSpaces(8)).append("YYYYYYYYYYYY").append(weekSchedule.toString()).append(getSpaces(47)).append("YA").append("\n");
		return finalWeekSchedule.toString().toCharArray();
	}

	private static char[] getMonthSchedule(List<Integer> schedules) {
		char[] monthDays = new char[31];
		StringBuffer sb1 = new StringBuffer();
		StringBuffer sb = new StringBuffer();
		Arrays.fill(monthDays, ' ');
		for (Integer schedule : schedules) {
			int day = schedule;//Integer.parseInt(schedule);
			monthDays[day-1] = '{';
		}
		for(int i = 0;i<monthDays.length;i++) {
			sb.append(monthDays[i]);
		}
		//System.out.println(sb.toString());
		sb1.append("M*M").append(getSpaces(8)).append("YYYYYYYYYYYY").append(sb.toString()).append(getSpaces(16)).append("305").append(getSpaces(6)).append("\n");
		sb1.append("V").append(getSpaces(1)).append("W").append(getSpaces(8)).append("YYYYYYYYYYYY").append(getDefaultSchedule(7, "")).append(getSpaces(47)).append("YA").append("\n");
		return sb1.toString().toCharArray();
	}

	private static char[] getNulCharacter(String scheduleDescription) {
		StringBuffer sb = new StringBuffer();
		// if schedule is blank then use default schedule
		sb.append("M*M").append(getSpaces(8)).append("YYYYYYYYYYYY").append(getDefaultSchedule(31, scheduleDescription)).append(getSpaces(16)).append("305").append(getSpaces(6)).append("\n");
		sb.append("V").append(getSpaces(1)).append("W").append(getSpaces(8)).append("YYYYYYYYYYYY").append(getDefaultSchedule(7, scheduleDescription)).append(getSpaces(47)).append("YA").append("\n");		
		return sb.toString().toCharArray();
	}

	private static char[] getDefaultSchedule(int i, String scheduleDescription) {
		StringBuffer sb = new StringBuffer();
		for (int j = 1; j <= i; j++) {
			sb.append("{");
		}
		return sb.toString().toCharArray();
	}
	
	private static char[] getDefaultSchedule(int i) {
		StringBuffer sb = new StringBuffer();
		for (int j = 1; j <= i; j++) {
			sb.append("{");
		}
		return sb.toString().toCharArray();
	}


	private static char[] writeClassBlock(String className, String jobName) {
		StringBuffer sb = new StringBuffer();		
		if (className.equals("")) {
			sb.append("NCLASS").append(getSpaces(73)).append("\n");
		} else if(!className.equals("")) {
			sb.append("NCLASS").append(getSpaces(27)).append(className).append(getSpaces(45)).append("\n");
		}		
		
		int totalSize = jobName.toCharArray().length;
		sb.append("W").append(getSpaces(78)).append("\n")
		.append("Y").append(getSpaces(78)).append("\n")
		.append("TNAME").append(getSpaces(3)).append(jobName)
				.append(getSpaces(65-totalSize)).append("X").append(getSpaces(5)).append("\n")
				.append("TUSER").append(getSpaces(5)).append("OPSADMIN").append(getSpaces(27)).append("A").append(getSpaces(33)).append("\n");
		//System.out.println(sb.toString());
		return sb.toString().toCharArray();
	}

	private static String getSpaces(int i) {
		String space = " ";
		StringBuffer sb = new StringBuffer();
		for (int k = 1; k <= i; k++) {
			sb.append(space);
		}
		return sb.toString();
	}

	private static char[] writeDSNBlock(JobDetails jobDetails) {
		StringBuffer sb = new StringBuffer();
		sb.append("NDSN").append(getSpaces(30)).append("LAST=Y,").append("PREFIX=CTD.SYS,");
		if (jobDetails.getStepName().equals("")) {
			sb.append(getSpaces(23)).append("\n");
		} else {
			int totalSize = 9+jobDetails.getStepName().length();
			sb.append("PGMSTEP=").append(jobDetails.getStepName() + ",").append(getSpaces(23-totalSize)).append("\n");
		}
		if (jobDetails.getProcStep().equals("") && !jobDetails.getDDName().equals("")) {
			int totalSize = 8 + jobDetails.getDDName().length();
			sb.append("XDDNAME=").append(jobDetails.getDDName()).append(getSpaces(79 - totalSize)).append("\n");
		} else if(!jobDetails.getProcStep().equals("") && jobDetails.getDDName().equals("")) {
			int totalSize = 10 + jobDetails.getProcStep().length();
			sb.append("XPROCSTEP=").append(jobDetails.getProcStep()).append(getSpaces(79 - totalSize)).append("\n");
		}
		else if (jobDetails.getProcStep().equals("") && jobDetails.getDDName().equals("")) {
			//System.out.println("DD name and proc step are blank....................");
		}
		else if(!jobDetails.getProcStep().equals("") && !jobDetails.getDDName().equals("")) {
			int totalSize = 18+ jobDetails.getProcStep().length() + jobDetails.getDDName().length();
			sb.append("XPROCSTEP=").append(jobDetails.getProcStep()).append(",");
			sb.append("DDNAME=").append(jobDetails.getDDName()).append(getSpaces(79 - totalSize)).append("\n");
		}
		int totalSize = jobDetails.getReportID().toCharArray().length;
		sb.append("W").append(getSpaces(78)).append("\n")
		.append("Y").append(getSpaces(78)).append("\n")
		.append("TNAME").append(getSpaces(3)).append(jobDetails.getReportID()).append(getSpaces(65-totalSize)).append("X").append(getSpaces(5)).append("\n");
		if(!jobDetails.getDistID().isEmpty()) {
			//System.out.println("Dist id list is not empty.....");
			for (String distID : jobDetails.getDistID()) {
				//System.out.println("dist id may be empty....");
				int totalSize1 = distID.toCharArray().length;
				sb.append("TUSER").append(getSpaces(5)).append(distID).append(getSpaces(35-totalSize1)).append("A").append(getSpaces(33)).append("\n");
			}			
		}
		else {
			//System.out.println("Dist id list is empty.....");
		}
		String str = "TBACKUP  BKP0031D";
		int len = str.length();
		sb.append("TBACKUP").append(getSpaces(2)).append("BKP0031D").append(getSpaces(80-len-16));
		//System.out.println(sb.toString());
		return sb.toString().toCharArray();
	}
}


	@Getter
	@Setter
	@ToString
	class JobDetails {
		private String reportID;
		private String stepName;
		private String procStep;
		private String DDName;
		private String className;
		private String form;
		private List<String> distID;
		private String category;
		private String scheduleDescription;	
	}

