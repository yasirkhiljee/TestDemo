package com.asadee.myapplication

import android.Manifest
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.view.View
import android.widget.EditText


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.Exception

import android.content.pm.PackageManager

import androidx.core.app.ActivityCompat





class MainActivity : AppCompatActivity() {

    private var editTextInput: EditText? = null
    private var filePath: File? = null


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        ActivityCompat.requestPermissions(
            this, arrayOf(
                Manifest.permission.READ_EXTERNAL_STORAGE,
                Manifest.permission.WRITE_EXTERNAL_STORAGE
            ),
            PackageManager.PERMISSION_GRANTED
        )

        editTextInput = findViewById(R.id.editTextTextPersonName)
        filePath = File(getExternalFilesDir(null), "Test.docx")

        try {
            if (!filePath!!.exists()) {
                filePath!!.createNewFile()
            }
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }


    fun buttonCreate(view: View) {
        try {
            val xwpfDocument = XWPFDocument()
            val xwpfParagraph = xwpfDocument.createParagraph()
            val xwpfRun = xwpfParagraph.createRun()
            xwpfRun.setText(editTextInput!!.text.toString())
            xwpfRun.fontSize = 24
            val fileOutputStream = FileOutputStream(filePath)
            xwpfDocument.write(fileOutputStream)
            if (fileOutputStream != null) {
                fileOutputStream.flush()
                fileOutputStream.close()
            }
            xwpfDocument.close()
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }
}