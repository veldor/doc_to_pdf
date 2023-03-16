@file:Suppress("HttpUrlsUsage")

package net.veldor.doc_to_pdf.model.handler

import net.veldor.doc_to_pdf.model.utils.ExecutionDateChecker
import org.apache.poi.xwpf.extractor.XWPFWordExtractor
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.xmlbeans.SimpleValue
import java.io.BufferedInputStream
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

class DocxDateRepairHandler {
    fun repairData(file: File): File? {
        try {
            FileInputStream(file).use { fis ->
                BufferedInputStream(fis).use { inputStream ->
                    XWPFDocument(inputStream).use { document ->
                        XWPFWordExtractor(document).use { extractor ->
                            val value = extractor.text
                            val executionDate: String? = ExecutionDateChecker.findDate(value)
                            if (executionDate != null) {
                                var paragraphForDelete: XWPFParagraph? = null
                                var titleForDelete: XWPFParagraph? = null
                                for (paragraph in document.paragraphs) {
                                    if (paragraph.text == """ (
Россия, 603001, Нижний Новгород
Нижневолжская наб., 4
ул. Родионова, 190 
(больница им. Семашко, 2-ой приемный покой)
ул. Советская, 12 (пл. Ленина)
(831) 461-82-83, 435-81-81
clinica@rdc.nnov.ru
www.мрт-кт
.р
ф
)"""
                                    ) {
                                        paragraphForDelete = paragraph
                                    } else if (paragraphForDelete != null && titleForDelete == null && paragraph.text == "Магнитно-резонансная томография") {
                                        titleForDelete = paragraph
                                    } else if (paragraphForDelete != null && paragraph.text == """ (
Общество с ограниченной ответственностью 
«Региональный диагностический центр» 
 
603001, Нижний Новгород, Нижне-Волжская набережная, 4 
(831) 461-82-83, 461-82-86)"""
                                    ) {
                                        titleForDelete = paragraph
                                    } else {
                                        for (run in paragraph.runs) {
                                            var cursor = run.ctr.newCursor()
                                            // получу начало объекта через курсор
                                            cursor.selectPath(
                                                "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:fldChar/@w:fldCharType"
                                            )
                                            while (cursor.hasNextSelection()) {
                                                cursor.toNextSelection()
                                                var obj = cursor.getObject()
                                                if ("begin" == (obj as SimpleValue).stringValue) {
                                                    // если элемент начинается с "begin"
                                                    // тут ожидается элементы формы, который мы будем использовать
                                                    cursor.toParent()
                                                    obj = cursor.getObject()
                                                    if (obj != null) {
                                                        var forms =
                                                            obj.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:ffData/w:name/@w:val")
                                                        if (forms.size > 0) {
                                                            val formName =
                                                                (forms[0] as SimpleValue).stringValue
                                                            if ("r_inspdate" == formName) {
                                                                // найдена форма, найду значение типа формы
                                                                forms =
                                                                    obj.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:ffData/w:textInput/w:type/@w:val")
                                                                if (forms.size > 0) {
                                                                    val currentType =
                                                                        (forms[0] as SimpleValue).stringValue
                                                                    if ("currentDate" == currentType || "currentTime" == currentType) {
                                                                        (forms[0] as SimpleValue).stringValue = "date"
                                                                        // Добавлю значение даты по умолчанию
                                                                        forms =
                                                                            obj.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:ffData/w:textInput")
                                                                        if (forms.size > 0) {
                                                                            val cursor1 = run.ctr.newCursor()
                                                                            // получу начало объекта через курсор
                                                                            cursor1.selectPath(
                                                                                "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:fldChar/w:ffData/w:textInput/w:type"
                                                                            )
                                                                            if (cursor1.hasNextSelection()) {
                                                                                cursor1.toNextSelection()
                                                                                cursor1.insertElement(
                                                                                    "default",
                                                                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                                                                )
                                                                                val cursor2 =
                                                                                    run.ctr.newCursor()
                                                                                cursor2.selectPath(
                                                                                    "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:fldChar/w:ffData/w:textInput/w:default"
                                                                                )
                                                                                cursor2.toNextSelection()
                                                                                cursor2.toEndToken()
                                                                                cursor2.insertAttributeWithValue(
                                                                                    "val",
                                                                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                                                                                    executionDate
                                                                                )
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            cursor = run.ctr.newCursor()
                                            // попробую найти форму с ФИО врача
                                            cursor.selectPath(
                                                "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:ddList"
                                            )
                                            if (cursor.hasNextSelection()) {
                                                cursor.toNextSelection()
                                                val obj = cursor.getObject()
                                                val result =
                                                    obj.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:result/@w:val")
                                                if (result.size == 1) {
                                                    val selectedDoctor =
                                                        (result[0] as SimpleValue).stringValue
                                                    if (!selectedDoctor.isEmpty()) {
                                                        // look for options list
                                                        val options =
                                                            obj.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:listEntry")
                                                        // проверка на то, что в списке именно врачи
                                                        if (options.size > 5 && selectedDoctor.toInt() <= options.size) {
                                                            val selectedOption = options[selectedDoctor.toInt()]
                                                            val nameAttribute = selectedOption.selectAttribute(
                                                                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                                                                "val"
                                                            )
                                                            if (nameAttribute != null) {
                                                                val docName =
                                                                    (nameAttribute as SimpleValue).stringValue
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (paragraphForDelete != null) {
                                    deleteParagraph(paragraphForDelete)
                                }
                                if (titleForDelete != null) {
                                    deleteParagraph(titleForDelete)
                                }
                                val newName: String
                                newName = if (file.name.length > 50) {
                                    file.name.substring(0, 50)
                                } else {
                                    file.name
                                }
                                val newFile = File.createTempFile("new_$newName", ".docx")
                                newFile.deleteOnExit()
                                document.isTrackRevisions = false
                                val fos = FileOutputStream(newFile)
                                document.write(fos)
                                fos.close()
                                fis.close()
                                inputStream.close()
                                return newFile
                            }
                        }
                    }
                }
            }
        } catch (e: Exception) {
            if (!e.message!!.contains("Zip bomb detected")) {
                e.printStackTrace()
            }
        }
        return null
    }

    private fun deleteParagraph(p: XWPFParagraph) {
        val doc = p.document
        val pPos = doc.getPosOfParagraph(p)
        doc.removeBodyElement(pPos)
    }
}