export interface IEventOccurrence {
  date: Date;
  title: string;
  eventType: string;
  imageUrl?: string;
  notes?: string;
  originalEventDate: Date;
}

export class RecurrenceCalculator {
  /**
   * Calculate the next occurrence of an event based on recurrence rules
   */
  public static calculateNextOccurrences(
    eventDate: Date,
    isAnnualRecurrence: boolean,
    recurrenceRule: string | null,
    title: string,
    eventType: string,
    imageUrl?: string,
    notes?: string,
    maxDays: number = 365
  ): IEventOccurrence[] {
    const occurrences: IEventOccurrence[] = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const endDate = new Date(today.getTime());
    endDate.setDate(endDate.getDate() + maxDays);

    if (recurrenceRule) {
      // Handle complex recurrence rules (e.g., Labor Day: first Monday of September)
      const ruleOccurrences = this.calculateByRecurrenceRule(
        eventDate,
        recurrenceRule,
        title,
        eventType,
        imageUrl,
        notes,
        today,
        endDate
      );
      occurrences.push(...ruleOccurrences);
    } else if (isAnnualRecurrence) {
      // Handle annual recurrence (birthdays, fixed holidays)
      const annualOccurrences = this.calculateAnnualOccurrences(
        eventDate,
        title,
        eventType,
        imageUrl,
        notes,
        today,
        endDate
      );
      occurrences.push(...annualOccurrences);
    } else {
      // Single occurrence (non-recurring event)
      const eventDateOnly = new Date(eventDate.getTime());
      eventDateOnly.setHours(0, 0, 0, 0);
      
      if (eventDateOnly >= today && eventDateOnly <= endDate) {
        occurrences.push({
          date: eventDateOnly,
          title,
          eventType,
          imageUrl,
          notes,
          originalEventDate: new Date(eventDate.getTime())
        });
      }
    }

    return occurrences;
  }

  /**
   * Calculate occurrences for a specific year (used for Holiday Quick View)
   */
  public static calculateOccurrencesForYear(
    eventDate: Date,
    isAnnualRecurrence: boolean,
    recurrenceRule: string | null,
    title: string,
    eventType: string,
    imageUrl?: string,
    notes?: string,
    year?: number
  ): IEventOccurrence[] {
    const targetYear = year || new Date().getFullYear();
    const startDate = new Date(targetYear, 0, 1); // January 1
    const endDate = new Date(targetYear, 11, 31); // December 31

    if (recurrenceRule) {
      return this.calculateByRecurrenceRule(
        eventDate,
        recurrenceRule,
        title,
        eventType,
        imageUrl,
        notes,
        startDate,
        endDate
      );
    } else if (isAnnualRecurrence) {
      return this.calculateAnnualOccurrences(
        eventDate,
        title,
        eventType,
        imageUrl,
        notes,
        startDate,
        endDate
      );
    } else {
      // Single occurrence - check if it falls in the target year
      const eventDateOnly = new Date(eventDate.getTime());
      eventDateOnly.setHours(0, 0, 0, 0);
      
      if (eventDateOnly >= startDate && eventDateOnly <= endDate) {
        return [{
          date: eventDateOnly,
          title,
          eventType,
          imageUrl,
          notes,
          originalEventDate: new Date(eventDate.getTime())
        }];
      }
      return [];
    }
  }

  /**
   * Calculate annual occurrences (same month/day each year)
   */
  private static calculateAnnualOccurrences(
    originalDate: Date,
    title: string,
    eventType: string,
    imageUrl: string | undefined,
    notes: string | undefined,
    startDate: Date,
    endDate: Date
  ): IEventOccurrence[] {
    const occurrences: IEventOccurrence[] = [];
    const originalMonth = originalDate.getMonth();
    const originalDay = originalDate.getDate();
    
    // Calculate occurrences for current year
    const currentYear = startDate.getFullYear();
    let checkDate = new Date(currentYear, originalMonth, originalDay);
    
    if (checkDate < startDate) {
      // If the date has passed this year, check next year
      checkDate = new Date(currentYear + 1, originalMonth, originalDay);
    }

    while (checkDate <= endDate) {
      occurrences.push({
        date: new Date(checkDate.getTime()),
        title,
        eventType,
        imageUrl,
        notes,
        originalEventDate: new Date(originalDate.getTime())
      });
      
      checkDate = new Date(checkDate.getFullYear() + 1, originalMonth, originalDay);
    }

    return occurrences;
  }

  /**
   * Calculate occurrences based on recurrence rule patterns
   * Supports patterns like: "MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1" (1st Monday of September)
   */
  private static calculateByRecurrenceRule(
    originalDate: Date,
    recurrenceRule: string,
    title: string,
    eventType: string,
    imageUrl: string | undefined,
    notes: string | undefined,
    startDate: Date,
    endDate: Date
  ): IEventOccurrence[] {
    const occurrences: IEventOccurrence[] = [];
    
    // Parse pattern: MONTHLY_BY_NTH_WEEKDAY:MM:WEEKDAY:NTH
    const pattern = /MONTHLY_BY_NTH_WEEKDAY:(\d{2}):(\w+):(\d+)/i;
    const match = recurrenceRule.match(pattern);
    
    if (match) {
      const month = parseInt(match[1], 10) - 1; // 0-indexed month
      const weekdayStr = match[2].toUpperCase();
      const nth = parseInt(match[3], 10);
      
      const weekdayMap: { [key: string]: number } = {
        'SUNDAY': 0,
        'MONDAY': 1,
        'TUESDAY': 2,
        'WEDNESDAY': 3,
        'THURSDAY': 4,
        'FRIDAY': 5,
        'SATURDAY': 6
      };
      
      const targetWeekday = weekdayMap[weekdayStr];
      
      if (targetWeekday !== undefined && month >= 0 && month <= 11) {
        let currentYear = startDate.getFullYear();
        const endYear = endDate.getFullYear();
        
        while (currentYear <= endYear) {
          const occurrence = this.findNthWeekdayInMonth(currentYear, month, targetWeekday, nth);
          
          if (occurrence >= startDate && occurrence <= endDate) {
            occurrences.push({
              date: new Date(occurrence.getTime()),
              title,
              eventType,
              imageUrl,
              notes,
              originalEventDate: new Date(originalDate.getTime())
            });
          }
          
          currentYear++;
        }
      }
    }
    
    return occurrences;
  }

  /**
   * Find the nth occurrence of a weekday in a given month/year
   * e.g., 1st Monday of September 2024, or last Monday of May (nth = 5)
   */
  private static findNthWeekdayInMonth(
    year: number,
    month: number,
    weekday: number,
    nth: number
  ): Date {
    // Handle "last" occurrence (nth = 5)
    if (nth === 5) {
      // Find the last occurrence by iterating backwards from the last day of the month
      // This is more reliable than calculating days to subtract
      const lastDay = new Date(year, month + 1, 0); // Day 0 of next month = last day of current month
      const lastDayDate = lastDay.getDate(); // Day number (1-31)
      
      // Iterate backwards from the last day to find the last occurrence of the target weekday
      for (let day = lastDayDate; day >= 1; day--) {
        const checkDate = new Date(year, month, day);
        if (checkDate.getDay() === weekday) {
          return checkDate;
        }
      }
      
      // This should never happen, but return last day as fallback
      return lastDay;
    }
    
    // Handle 1st-4th occurrences
    // Start with the first day of the month
    const firstDay = new Date(year, month, 1);
    const firstDayWeekday = firstDay.getDay();
    
    // Calculate days to add to reach the target weekday
    let daysToAdd = (weekday - firstDayWeekday + 7) % 7;
    
    // Add weeks for nth occurrence (e.g., for 2nd occurrence, add 7 more days)
    daysToAdd += (nth - 1) * 7;
    
    const targetDate = new Date(year, month, 1 + daysToAdd);
    
    return targetDate;
  }
}
