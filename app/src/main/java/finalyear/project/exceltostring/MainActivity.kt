package finalyear.project.exceltostring

import android.content.Context
import android.os.Bundle
import android.util.Log
import android.widget.Button
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import com.orhanobut.logger.AndroidLogAdapter
import com.orhanobut.logger.Logger
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.FileOutputStream
import java.io.InputStreamReader
import java.io.OutputStreamWriter
import java.lang.StringBuilder


class MainActivity : AppCompatActivity() {

    private val TAG = "MainActivity"
    private val FILE_NAME = "notification_string.txt"
    private lateinit var textView: TextView
    private lateinit var showButton: Button


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)


        Logger.addLogAdapter(AndroidLogAdapter())


        textView = findViewById(R.id.hadithTextView)
        showButton = findViewById(R.id.show)

        showButton.setOnClickListener {
            showHadith()
        }

        readNotificationContent()


    }

    // Reading from excel file
    private fun readNotificationContent() {

        try {
            val inputStream = resources.openRawResource(R.raw.notification_content)

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

            saveNotification(hadiths)
        } catch (e: Exception) {
            Log.d(TAG, "error: ${e.message}")
        }
    }

    // Writing into text file
    private fun saveNotification(hadiths: ArrayList<String>) {
        try {

            var fileOutputStream = OutputStreamWriter(openFileOutput(FILE_NAME, Context.MODE_PRIVATE))


            hadiths.forEach {
                fileOutputStream.append(it)
                fileOutputStream.append("\n\n\n")
            }
            fileOutputStream.close()

        } catch (e: Exception) {
            Log.d(TAG, "FileWriteError: ${e.message}")
        }
    }

    // Reading from text file
    private fun showHadith() {
        try {
            val inputStream = openFileInput(FILE_NAME)

            if (inputStream != null) {
                val inputStreamReader = InputStreamReader(inputStream)
                val bufferReader = BufferedReader(inputStreamReader)
                var value = bufferReader.readLine()
                var stringBuilder = StringBuilder()

                while (value != null) {
                    stringBuilder.append("\n").append(value)
                    value = bufferReader.readLine()
                }

                inputStream.close()

                textView.setText(stringBuilder.toString())
            }
        }catch (e: Exception) {
            Log.d(TAG, "FileReadError: ${e.message}")
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
}
