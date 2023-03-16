package net.veldor.doc_to_pdf.model.utils

import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFTable
import java.io.BufferedInputStream
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

class DocCleaner {
    fun clean(file: File) {
        try {
            FileInputStream(file).use { fis ->
                BufferedInputStream(fis).use { inputStream ->
                    XWPFDocument(inputStream).use { document ->
                        var text: String
                        var target: XWPFParagraph? = null
                        var next: XWPFParagraph? = null
                        var paragraphForDelete: XWPFParagraph? = null
                        var titleForDelete: XWPFParagraph? = null
                        var tableForDelete: XWPFTable? = null
                        for (table in document.tables) {
                            if (table.text == """Общество с ограниченной ответственностью	«Региональный диагностический центр»	
603001, Нижний Новгород, Нижне-Волжская набережная, 4	(831) 461-82-83, 461-82-86
"""
                            ) {
                                tableForDelete = table
                            }
                        }
                        if (tableForDelete != null) {
                            val posOfTable = document.getPosOfTable(tableForDelete);
                            document.removeBodyElement(posOfTable)
                        }
                        for (paragraph in document.paragraphs) {
                            text = paragraph.text
                            if (text.contains("Evaluation Only. Created with Aspose.Words.")) {
                                target = paragraph
                            } else if (text == """ (
Россия, 603001, Нижний Новгород
Нижневолжская наб., 4
ул. Родионова, 190 
(больница им. Семашко, 2-ой приемный покой)
ул. Советская, 12 (пл. Ленина)
(831) 461-82-83, 435-81-81
clinica@rdc.nnov.ru
www.мрт-кт.рф
)"""
                            ) {
                                paragraphForDelete = paragraph
                            } else if (paragraphForDelete != null && titleForDelete == null && text == "Магнитно-резонансная томография") {
                                titleForDelete = paragraph
                            }
                            if (target != null && next == null) {
                                next = paragraph
                            }
                        }
                        if (target != null) {
                            document.removeBodyElement(document.getPosOfParagraph(target))
                            document.removeBodyElement(document.getPosOfParagraph(next))
                            for (header in document.headerList) {
                                header.setHeaderFooter(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr.Factory.newInstance())
                            }
                            for (footer in document.footerList) {
                                footer.setHeaderFooter(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr.Factory.newInstance())
                            }
                        }
                        if (paragraphForDelete != null) {
                            document.removeBodyElement(document.getPosOfParagraph(paragraphForDelete))
                            document.removeBodyElement(document.getPosOfParagraph(titleForDelete))
                        }
                        document.isTrackRevisions = false
                        val fos = FileOutputStream(file)
                        document.write(fos)
                        fos.close()
                    }
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }
}