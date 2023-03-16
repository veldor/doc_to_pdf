package net.veldor.doc_to_pdf.model.handler

import java.io.File
import java.util.*


class QueueHandler {
    private lateinit var queue: Queue<File>
    private var counter = 0

    private fun recursiveAddToQueue(entity: File) {
        // проверю содержимое папки
        val files = entity.listFiles()
        if (files != null) {
            if (files.isNotEmpty()) {
                files.forEach {
                    if (it.isDirectory) {
                        recursiveAddToQueue(it)
                    } else {
                        if (!it.name.startsWith("~$") && FileHandler.accept(it)) {
                            queue.add(it)
                            print("\rQueued $counter")
                            counter++
                        }
                    }
                }
            }
        }
    }

    fun fillQueue(source: File): Queue<File> {
        queue = LinkedList()
        recursiveAddToQueue(source)
        return queue
    }
}