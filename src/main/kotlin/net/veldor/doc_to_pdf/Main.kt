import net.veldor.doc_to_pdf.model.converter.PdfConverter
import net.veldor.doc_to_pdf.model.handler.DocDateRepairHandler
import net.veldor.doc_to_pdf.model.handler.DocxDateRepairHandler
import net.veldor.doc_to_pdf.model.handler.FileHandler
import net.veldor.doc_to_pdf.model.handler.QueueHandler
import net.veldor.doc_to_pdf.model.utils.RandomString
import java.io.File
import java.util.concurrent.TimeUnit

fun main(args: Array<String>) {

    if (args.size != 2) {
        println("Main 4 required 2 arguments: source dir and destination dir")
        return
    }
    val startTimestamp = System.currentTimeMillis()
    var timeLeft: Long
    var fileHandleTime: Long
    var timeToFinish: Long
    val source = File(args[0])
    val destination = File(args[1])
    if (!source.exists() || !source.isDirectory || !destination.exists() || !destination.isDirectory) {
        println("Main 12 wrong directories, check it!")
        return
    }
    val queue = QueueHandler().fillQueue(source)
    val fullLength = queue.size
    var currentFile: File
    var fixedFile: File?
    var tempPdfFile: File
    var destinationFile: File
    var handledCounter = 0
    while (queue.isNotEmpty()) {
        handledCounter++
        timeLeft = System.currentTimeMillis() - startTimestamp
        fileHandleTime = timeLeft / handledCounter
        timeToFinish = fileHandleTime * (fullLength - handledCounter)
        currentFile = queue.poll()
            print("\nОбрабатываю $handledCounter из $fullLength ${currentFile.name} (До завершения: ${TimeUnit.MILLISECONDS.toHours(timeToFinish)}:${TimeUnit.MILLISECONDS.toMinutes(timeToFinish / 60000)}:${TimeUnit.MILLISECONDS.toSeconds(timeToFinish / 60)})")
        destinationFile = FileHandler.getDestination(currentFile, destination)
        if (destinationFile.isFile && destinationFile.length() > 200) {
            // looks like file is converted, skip
            print("\r${currentFile.name} converted, skip")
            continue
        }
        tempPdfFile = File.createTempFile(RandomString.getRandomString(32), ".pdf")
        // fix document date
        // обработаю дату обследования
        var tryCounter = 0
        while (tryCounter < 6) {
            try {
                if (currentFile.extension == "doc") {
                    fixedFile = DocDateRepairHandler().repairData(currentFile)
                    PdfConverter().convertDoc(fixedFile, tempPdfFile)
                } else if (currentFile.extension == "docx") {
                    fixedFile = DocxDateRepairHandler().repairData(currentFile)
                    if (fixedFile != null) {
                        PdfConverter().convertDocx(fixedFile, tempPdfFile)
                    }
                }
                break
            } catch (t: Throwable) {
                t.printStackTrace()
                tryCounter++
            }
        }
        if (tempPdfFile.exists() && tempPdfFile.isFile && tempPdfFile.length() > 200) {
            // get name for file and move it to destination
            tempPdfFile.renameTo(destinationFile)
        }
        else{
            println("Main 75 error when handle ${currentFile.name}")
        }
    }
    PdfConverter.converter.shutDown()
}