package acambieri.calendarextractor
import com.google.api.client.util.DateTime
import com.google.api.services.calendar.Calendar
import com.google.api.services.calendar.model.Event
import org.apache.poi.ss.usermodel.*
import java.io.*
import java.nio.charset.Charset
import java.nio.file.Paths
import java.security.GeneralSecurityException
import java.text.SimpleDateFormat
import java.time.*
import java.util.*

object Rapportino {

    private val OUTPUT_FILE=Paths.get("out.xlsx")
    private val TEMPLATE_FILE= Paths.get("template.xlsx")
    private val config = ConfigProperties(Properties().apply {
        FileInputStream("config.properties").use { fis ->
            load(InputStreamReader(fis, Charset.forName("UTF-8")))
        }
    })

    @Throws(IOException::class, GeneralSecurityException::class)
    @JvmStatic
    fun main(args: Array<String>) {
        val year: String
        when(args.size){
            0 -> error("Usage: java -jar Rapportino.jar monthNumber [4digitYear]")
            1 -> year = GregorianCalendar.getInstance().get(java.util.Calendar.YEAR).toString()
            2 -> year = args[1]
            else -> error("Usage: java -jar Rapportino.jar monthNumber [4digitYear]")
        }
        val service = getCalendarService()
        val workbook = WorkbookFactory.create(TEMPLATE_FILE.toFile())
        val sheet = workbook.getSheetAt(0)
        for(cal in config.getCalendars()) {
            getEvents(service, args[0], year, cal).forEach { event ->
                val row = calculateDayStartRow(event)
                var found = false
                for (i in 0..config.maxJobPerDay - 1) {
                    val cRow = calculateCurrentRow(row, cal, i)
                    val descriptionCell = getDescriptionCell(i)
                    val hoursCell = getHoursCell(cal, i)
                    if (sheet.getRow(cRow)
                            .getCellValue(descriptionCell)
                            == null) {
                        found = true
                        sheet.getRow(cRow)
                                .getCell(descriptionCell)
                                ?.setCellValue(calculateCellValue(cal, event))
                        sheet.getRow(cRow)
                                .getCell(hoursCell)
                                ?.setCellValue(event.durata.toDouble())
                        break
                    }
                }
                if (!found) {
                    println("Skipped event for date ${event.data}: max jobs per day reached")
                }
            }
        }
        workbook.forceFormulaRecalculation = true
        FileOutputStream(OUTPUT_FILE.toFile()).use {
            workbook.write(it)
        }
    }

    private fun calculateDayStartRow(event: CalendarEvent): Int {
        return config.daysStartRow-1 + ((event.data.dayOfMonth-1)*config.nextDayRowDistance)
    }

    private fun getDescriptionCell(offset: Int): Int {
        return config.descriptionColDistance+(config.nextJobColDistance*offset)
    }

    private fun calculateCellValue(calendar: String, event: CalendarEvent): String {
        return when(calendar){
            config.WORKTIME_CALENDAR_NAME -> event.nome
            config.PUBLIC_HOLIDAY_CALENDAR_NAME -> config.publicHolidayDescription
            config.HOLIDAY_CALENDAR -> config.holidayDescription
            else -> error("Unsupported calendar")
        }
    }

    private fun getHoursCell(calendar: String, offset: Int): Int {
        return when(calendar){
            config.WORKTIME_CALENDAR_NAME -> config.worktimeColDistance
            config.PUBLIC_HOLIDAY_CALENDAR_NAME -> config.publicHolidayColDistance
            config.HOLIDAY_CALENDAR -> config.holidayColDistance
            else -> error("Unsupported calendar")
        }+(config.nextJobColDistance*offset)
    }

    private fun calculateCurrentRow (row: Int,calendar : String,offset: Int) : Int {
        return (row+
                when(calendar){
                    config.WORKTIME_CALENDAR_NAME -> config.descriptionRowDistance
                    config.PUBLIC_HOLIDAY_CALENDAR_NAME -> config.publicHolidayRowDistance
                    config.HOLIDAY_CALENDAR -> config.holydayRowDistance
                    else -> error("Unsupported calendar")
                }
        +(config.nextJobRowDistance*offset))
    }

    private fun getEvents(service: Calendar, month: String,year: String,calendar: String): List<CalendarEvent> {
        val df = SimpleDateFormat("dd/MM/yyyy")
        df.isLenient=false
        val minTime = DateTime(df.parse("1/${month}/${year}"))
        val maxTime = DateTime(df.parse("${YearMonth.of(Integer.parseInt(year),Integer.parseInt(month)).lengthOfMonth()}/${month}/${year}"))
        val calendars = service.CalendarList().list().execute();
        val calendarsFiltered = calendars.items.filter { it -> it.summary == calendar }
        if(calendarsFiltered.size > 0) {
            val commesseCalendar = calendarsFiltered.first()
            val events = with(service.events().list(commesseCalendar.id)) {
                timeMin = minTime
                timeMax = maxTime
                orderBy = "startTime"
                singleEvents = true
                execute()
            }
            println("Calendar ${calendar} has ${events.items.size} events")
            return events.items.map { event -> eventToCommessa(event) }
        }
        else{
            println("Calendar ${calendar} not found! skipping...")
            return ArrayList(0)
        }
    }

    private fun eventToCommessa(event:Event) : CalendarEvent {
        return CalendarEvent(
                nome= event.summary,
                durata=when{
                    event.isEndTimeUnspecified || event.start.dateTime == null -> 8
                    event.start.dateTime != null && event.end.dateTime != null -> Math.round(Duration.between(LocalDateTime.ofEpochSecond(event.start.dateTime.value/1000,0, ZoneOffset.UTC),
                            LocalDateTime.ofEpochSecond(event.end.dateTime.value/1000,0, ZoneOffset.UTC)).toMinutes()/60.0).toInt()
                    else -> 0
                },
                data=LocalDateTime.ofEpochSecond(
                        when {
                                        event.start.dateTime == null -> event.start.date
                                        else -> event.start.dateTime
                                    }.value/1000,
                        0,
                        ZoneOffset.UTC).toLocalDate()
        )
    }
}

data class CalendarEvent(val nome:String, val durata: Int, val data: LocalDate) {

}


fun Row.getCellValue(cellNum: Int): String?{
    val cell: Cell? = getCell(cellNum)
    when(cell?.cellType){
        Cell.CELL_TYPE_BLANK -> return null;
        Cell.CELL_TYPE_NUMERIC -> return cell.numericCellValue.toString()
        Cell.CELL_TYPE_STRING -> return if(cell.stringCellValue == null || cell.stringCellValue.isEmpty() ) null else cell.stringCellValue
        else -> return null;
    }
}