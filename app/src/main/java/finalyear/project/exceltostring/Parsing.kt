package finalyear.project.exceltostring

import android.content.Context
import android.util.Log
import com.loopj.android.http.AsyncHttpClient
import com.loopj.android.http.FileAsyncHttpResponseHandler
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.FileInputStream
import java.io.InputStreamReader
import java.io.OutputStreamWriter
import java.lang.StringBuilder

private val TAG = "MainActivity"
private val FILE_NAME = "notification_string.txt"
private val URL = "https://github.com/Shuvo1260/ExcelToString/blob/master/app/src/main/res/raw/notification_content.xlsx?raw=true"

fun main() {
    readNotificationContent()
}


// Reading from excel file
private fun readNotificationContent() {

    try {
        val inputStream = FileInputStream("./notification_content.xlsx")

        val workBook = XSSFWorkbook(inputStream)

        val sheet = workBook.getSheetAt(0)

        val rowCount = sheet.physicalNumberOfRows

        val formulaEvaluator = workBook.creationHelper.createFormulaEvaluator()

        var hadiths = arrayListOf<String>()

        for (rowIndex in 1 until rowCount-1) {
            val row = sheet.getRow(rowIndex)

            val cellsCount = row.physicalNumberOfCells

            var cellValues = arrayListOf<String>()

            for (cellIndex in 0 until cellsCount) {
                val value = getCellAsString(row, cellIndex, formulaEvaluator)

                cellValues.add(value)

            }

            // Formating
            val content = cellValues[0].trim() +
                    "\n\nরেফারেন্সঃ\n" + cellValues[1].trim() +
                    "\n" + cellValues[2].trim()

            hadiths.add(content)

            Log.d(TAG, "Hadith: $content")


        }

//            showHadith(hadiths[1])

        println("${hadiths[0]}")
//        saveNotification(hadiths)
    } catch (e: Exception) {
        Log.d(TAG, "error: ${e.message}")
    }
}

private fun getCellAsString(row: Row, cellIndex: Int, formulaEvaluator: FormulaEvaluator): String {
    var value = ""
    try {
        val cell = row.getCell(cellIndex)
        val cellValue = formulaEvaluator.evaluate(cell)

        value += " " + cellValue.stringValue
    }catch (e: Exception) {
        Log.d(TAG, "FormatingError: ${e.message}")
    }

    return value
}

//// Writing into text file
//private fun saveNotification(hadiths: ArrayList<String>) {
//    try {
//
//        var fileOutputStream = OutputStreamWriter(openFileOutput(FILE_NAME, Context.MODE_PRIVATE))
//
//
//        hadiths.forEach {
//            fileOutputStream.append(it)
//            fileOutputStream.append("\n\n\n")
//        }
//        fileOutputStream.close()
//
//    } catch (e: Exception) {
//        Log.d(TAG, "FileWriteError: ${e.message}")
//    }
//}
//
//// Reading from text file
//private fun showHadith() {
//    try {
//        val inputStream = openFileInput(FILE_NAME)
//
//        if (inputStream != null) {
//            val inputStreamReader = InputStreamReader(inputStream)
//            val bufferReader = BufferedReader(inputStreamReader)
//            var value = bufferReader.readLine()
//            var stringBuilder = StringBuilder()
//
//            while (value != null) {
//                stringBuilder.append("\n").append(value)
//                value = bufferReader.readLine()
//            }
//
//            inputStream.close()
//
//            textView.setText(stringBuilder.toString())
//        }
//    }catch (e: Exception) {
//        Log.d(TAG, "FileReadError: ${e.message}")
//    }
//}
