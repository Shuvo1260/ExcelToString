package finalyear.project.exceltostring

import android.os.Bundle
import android.util.Log
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import com.orhanobut.logger.AndroidLogAdapter
import com.orhanobut.logger.Logger
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook


class MainActivity : AppCompatActivity() {

    private val TAG = "MainActivity"

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)


        Logger.addLogAdapter(AndroidLogAdapter())

        readNotificationContent()
    }

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

//                    Log.d(TAG, "Value: $value")
//                    Logger.d(value)
                    cellValues.add(value)

                }

                // Formating
                val content = cellValues[0].trim() +
                        "\n\nরেফারেন্সঃ\n" + cellValues[1].trim() +
                        "\n" + cellValues[2].trim()

                hadiths.add(content)

                Log.d(TAG, "Hadith: $content")


            }

            showHadith(hadiths[1])
        } catch (e: Exception) {
            Log.d(TAG, "error: ${e.message}")
        }
    }

    private fun showHadith(hadith: String) {
        var hadithTextView: TextView = findViewById(R.id.hadithTextView)

        hadithTextView.text = hadith
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
