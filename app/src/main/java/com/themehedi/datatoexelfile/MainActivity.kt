package com.themehedi.datatoexelfile

import android.os.Build
import android.os.Bundle
import android.os.storage.StorageManager
import android.view.View
import android.widget.ArrayAdapter
import android.widget.ListView
import android.widget.Toast
import androidx.annotation.RequiresApi
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.File
import java.io.FileOutputStream


class MainActivity : AppCompatActivity() {
    private lateinit var listView: ListView
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        listView = findViewById(R.id.listView)
        val stringArray =
            arrayOf("This", "is", "an", "example", "Excel", "Sheet", "from", "programmer", "world")

        listView.adapter = ArrayAdapter(
            this@MainActivity,
            android.R.layout.simple_list_item_1,
            stringArray
        )
    }

    fun buttonCreateExcelFile(view: View?) {
        val hssfWorkbook = HSSFWorkbook()
        val hssfSheet = hssfWorkbook.createSheet("MySheet")
        val hssfRow = hssfSheet.createRow(0)
        for (i in 0 until listView.count) {
            val hssfCell = hssfRow.createCell(i)
            hssfCell.setCellValue(listView.getItemAtPosition(i).toString())
        }
        saveWorkBook(hssfWorkbook)
    }

    private fun saveWorkBook(hssfWorkbook: HSSFWorkbook) {
        val storageManager = getSystemService(STORAGE_SERVICE) as StorageManager
        val storageVolume = storageManager.storageVolumes[0] // internal storage
        val fileOutput = File(storageVolume.directory!!.path + "/Download/ProgrammerWorld.xls")
        try {
            val fileOutputStream = FileOutputStream(fileOutput)
            hssfWorkbook.write(fileOutputStream)
            fileOutputStream.close()
            hssfWorkbook.close()
            Toast.makeText(this, "File Created Successfully", Toast.LENGTH_LONG).show()
        } catch (e: Exception) {
            Toast.makeText(this, "File Creation Failed", Toast.LENGTH_LONG).show()
            throw RuntimeException(e)
        }
    }
}