package net.veldor.doc_to_pdf.model.handler

import com.aspose.words.*
import net.veldor.doc_to_pdf.model.utils.DocCleaner
import net.veldor.doc_to_pdf.model.utils.ExecutionDateChecker
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream

class DocDateRepairHandler {

    fun repairData(file: File): File {
        var formFound = false
        var form: FormField? = null
        val fis = FileInputStream(file)
        val doc = Document(fis)
        val s: Section = doc.firstSection
        val b: Body = s.body
        val ps: Array<Paragraph> = b.paragraphs.toArray()
        for (p in ps) {
            for (o in p.childNodes) {
                val currentNode: Node = o as Node
                if (formFound && currentNode.nodeType == NodeType.RUN) {
                    if (ExecutionDateChecker.isDate(currentNode.text)) {
                        form!!.textInputDefault = currentNode.text
                    }
                }
                if (currentNode.nodeType == NodeType.FORM_FIELD) {
                    form = o as FormField
                    if (form.name
                            .equals("r_inspdate") || form.textInputType == TextFormFieldType.CURRENT_TIME || form.textInputType == TextFormFieldType.CURRENT_DATE
                    ) {
                        form.textInputType = TextFormFieldType.DATE
                        formFound = true
                    }
                }
            }
        }
        val newName: String = if (file.name.length > 50) {
            file.name.substring(0, 50)
        } else {
            file.name
        }
        doc.acceptAllRevisions()
        val newFile = File.createTempFile("new_$newName", ".docx")
        val fos = FileOutputStream(newFile)
        doc.save(fos, SaveFormat.DOCX)
        fos.close()
        fis.close()
        fos.close()
        // теперь придётся убрать параграфы из нового документа
        DocCleaner().clean(newFile)
        return newFile
    }

}