package net.veldor.doc_to_pdf.model.handler

import java.io.File

object FileHandler {
    fun accept(it: File?): Boolean {
        return it?.extension == "doc" || it?.extension == "docx"
    }

    fun getDestination(currentFile: File, destinationDir: File): File {
        val name =
            if (currentFile.name.length > 100) {
                currentFile.name.substring(0, 100)
            } else {
                currentFile.name
            }
        val dir = name.substring(0, 2)
        val dd = File(destinationDir, dir)
        if (!dd.isDirectory) {
            dd.mkdirs()
        }
        return File(dd, "$name.pdf")
    }
}