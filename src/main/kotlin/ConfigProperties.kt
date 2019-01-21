package acambieri.calendarextractor
import java.util.*

data class ConfigProperties(val daysStartRow: Int,
                            val maxJobPerDay: Int,
                            val descriptionColDistance: Int,
                            val descriptionRowDistance: Int,
                            val worktimeColDistance: Int,
                            val worktimeRowDistance: Int,
                            val holidayColDistance: Int,
                            val holydayRowDistance: Int,
                            val publicHolidayColDistance: Int,
                            val publicHolidayRowDistance: Int,
                            val nextDayColDistance: Int,
                            val nextDayRowDistance: Int,
                            val nextJobColDistance: Int,
                            val nextJobRowDistance: Int,
                            val WORKTIME_CALENDAR_NAME : String,
                            val PUBLIC_HOLIDAY_CALENDAR_NAME: String,
                            val HOLIDAY_CALENDAR: String,
                            val publicHolidayDescription: String,
                            val holidayDescription: String
                            ){

    constructor(properties: Properties):
        this(properties["daysStartRow"].toString().toInt(),
                properties["maxJobPerDay"].toString().toInt(),
                properties["descriptionRowDistance"].toString().toInt(),
                properties["descriptionColDistance"].toString().toInt(),
                properties["worktimeColDistance"].toString().toInt(),
                properties["worktimeRowDistance"].toString().toInt(),
                properties["holidayColDistance"].toString().toInt(),
                properties["holydayRowDistance"].toString().toInt(),
                properties["publicHolidayColDistance"].toString().toInt(),
                properties["publicHolidayRowDistance"].toString().toInt(),
                properties["nextDayColDistance"].toString().toInt(),
                properties["nextDayRowDistance"].toString().toInt(),
                properties["nextJobColDistance"].toString().toInt(),
                properties["nextJobRowDistance"].toString().toInt(),
                properties["worktimeCalendar"].toString(),
                properties["publicHolidayCalendar"].toString(),
                properties["holidayCalendar"].toString(),
                properties["publicHolidayDescription"].toString(),
                properties["holidayDescription"].toString()
                )

    fun getCalendars() : List<String>{
        return Arrays.asList(WORKTIME_CALENDAR_NAME, PUBLIC_HOLIDAY_CALENDAR_NAME, HOLIDAY_CALENDAR)
    }


}

