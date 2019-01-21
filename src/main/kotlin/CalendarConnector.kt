package acambieri.calendarextractor
import com.google.api.client.auth.oauth2.Credential
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.jackson2.JacksonFactory
import com.google.api.client.util.store.FileDataStoreFactory
import com.google.api.services.calendar.Calendar
import com.google.api.services.calendar.CalendarScopes
import java.io.FileInputStream
import java.io.IOException
import java.io.InputStreamReader
import java.nio.file.Paths

val APPLICATION_NAME = "Rapportino"
val JSON_FACTORY = JacksonFactory.getDefaultInstance()
val TOKENS_DIRECTORY_PATH = Paths.get("${System.getProperty("user.home")}/.credentials/rapportino/tokens").toAbsolutePath()
val SCOPES = listOf(CalendarScopes.CALENDAR_READONLY)
val CREDENTIALS_FILE = Paths.get("${System.getProperty("user.home")}/.credentials/rapportino/credentials.json").toAbsolutePath()

@Throws(IOException::class)
fun getCredentials(HTTP_TRANSPORT: NetHttpTransport): Credential {
    //This function is taken from google examples
    val credentials = FileInputStream(CREDENTIALS_FILE.toFile())
    val clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, InputStreamReader(credentials))
    val flow = GoogleAuthorizationCodeFlow.Builder(
            HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
            .setDataStoreFactory(FileDataStoreFactory(TOKENS_DIRECTORY_PATH.toFile()))
            .setAccessType("offline")
            .build()
    val receiver = LocalServerReceiver.Builder().setPort(8888).build()
    return AuthorizationCodeInstalledApp(flow, receiver).authorize("user")
}

fun getCalendarService(): Calendar {
    //This function is taken from google examples
    val HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport()
    return Calendar.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
            .setApplicationName(APPLICATION_NAME)
            .build()
}