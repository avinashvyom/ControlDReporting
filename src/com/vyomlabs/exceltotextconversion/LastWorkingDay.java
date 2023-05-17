package com.vyomlabs.exceltotextconversion;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Month;
import java.util.ArrayList;
import java.util.List;

public class LastWorkingDay {
	
	int year = 2023;


//	public static void main(String[] args) {
//		int[] year = new int[] {2024};
//		for (int y : year) {
//			for (Month month : Month.values()) {
//				LocalDate ld = LocalDate.of(y, month, month.length(LocalDate.of(y, 1, 1).isLeapYear()));
//				while (ld.getDayOfWeek() == DayOfWeek.SATURDAY || ld.getDayOfWeek() == DayOfWeek.SUNDAY) {
//					ld = ld.minusDays(1);
//				}
//				//System.out.println("Last day of " + month + " " + y + " : " + ld + "DAY is: " + ld.getDayOfWeek());
//			}
//		}
//		LocalDate ld = LocalDate.of(2023, Month.APRIL, Month.APRIL.length(false));
//		//System.out.println("Last day of " + Month.APRIL + " " + 2023 + " : " + ld + " DAY is: " + ld.getDayOfWeek());
//		//System.out.println(6-5.7);
//		LastWorkingDay lastWorkingDay = new LastWorkingDay();
//		System.out.println("Last working is : "+lastWorkingDay.getLastWorkingDay(lastWorkingDay.getCurrentMonth()));
//		//lastWorkingDay.printAllWorkingDays(Month.APRIL, 2023);
//	}

	public List<String> getLastWorkingDay(Month month) {
		LocalDate ld = LocalDate.of(year, month, month.length(LocalDate.of(year, 1, 1).isLeapYear()));
		while (ld.getDayOfWeek() == DayOfWeek.SATURDAY || ld.getDayOfWeek() == DayOfWeek.SUNDAY) {
			ld = ld.minusDays(1);
		}
		Integer lastWorkingDayOfMonth = ld.getDayOfMonth();
		String dayOfWeek = ld.getDayOfWeek().name();
		System.out.println("Last working day : " + lastWorkingDayOfMonth + "Day of week : " + dayOfWeek);
		List<String> lastworkingDaydetails = new ArrayList<>();
		lastworkingDaydetails.add(lastWorkingDayOfMonth.toString());
		lastworkingDaydetails.add(dayOfWeek);
		return lastworkingDaydetails;
	}
	public Month getCurrentMonth() {
		LocalDate ld = LocalDate.ofYearDay(year, LocalDate.now().getDayOfYear());
		return ld.getMonth();
	}
	
//	public void printAllWorkingDays(Month month, int year) {
//		LocalDate ld = LocalDate.of(year, month, 1);
//		for (int i = 1; i <= month.maxLength(); i++) {
//			if (LocalDate.of(year, month, i).getDayOfWeek() == DayOfWeek.SATURDAY
//					|| LocalDate.of(year, month, i).getDayOfWeek() == DayOfWeek.SUNDAY) {
//				continue;
//			} else {
//				System.out.println(LocalDate.of(year, month, i) + ":" + LocalDate.of(year, month, i).getDayOfWeek());
//			}
//		}
//	}
}
