package net.veldor.doc_to_pdf.model.converter

import com.documents4j.api.DocumentType
import com.documents4j.api.IConverter
import com.documents4j.job.LocalConverter
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

class PdfConverter {

    companion object {
        var converter: IConverter = LocalConverter.builder().build()
    }


    fun closeConverter() {
        converter.shutDown()
    }

    fun convertDoc(source: File, destination: File): File? {
        // создам файл во временной папке
        var tryCounter = 1
        while (true) {
            // сначала попробую конвертацию с помощью MS Word
            try {
                FileInputStream(source).use { docxInputStream ->
                    FileOutputStream(destination).use { outputStream ->
                        converter.convert(docxInputStream).`as`(DocumentType.MS_WORD).to(outputStream)
                            .`as`(DocumentType.PDF)
                            .execute()
                        if (destination.length() > 0) {
                            return destination
                        }
                    }
                }
            } catch (e: Throwable) {
                e.printStackTrace()
                tryCounter++
                if (tryCounter > 5) {
                    println("FILE TOTALLY NOT CONVERTED")
                    break
                }
            }
        }
        return if (destination.length() > 0) {
            destination
        } else null
    }

    fun convertDocx(f: File, destination: File): File? {
        // создам файл во временной папке
        var tryCounter = 1
        // сначала попробую конвертацию с помощью MS Word
        while (true) {
            try {
                FileInputStream(f).use { docxInputStream ->
                    FileOutputStream(destination).use { outputStream ->
                        converter.convert(docxInputStream).`as`(DocumentType.DOCX).to(outputStream)
                            .`as`(DocumentType.PDF)
                            .execute()
                        if (destination.length() > 0) {
                            //converter.shutDown();
                            return destination
                        }
                    }
                }
            } catch (e: Throwable) {
                println(e.message)
                tryCounter++
                if (tryCounter > 5) {
                    break
                }
            }
        }
        return if (destination.length() > 0) {
            destination
        } else null
    }
}