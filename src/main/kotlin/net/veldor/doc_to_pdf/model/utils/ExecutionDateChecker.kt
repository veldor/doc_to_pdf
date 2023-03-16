package net.veldor.doc_to_pdf.model.utils

import java.util.*
import java.util.regex.Matcher
import java.util.regex.Pattern
import kotlin.collections.ArrayList

object ExecutionDateChecker {

    private val datePattern: Pattern = Pattern.compile(
        "^\\s*(\\d{1,2}[/|.]\\d{1,2}[/|.]\\d{2,4}).*",
        Pattern.CASE_INSENSITIVE or Pattern.UNICODE_CASE
    )

    fun isDate(text: String?): Boolean {
        val matcher: Matcher = datePattern.matcher(text!!.trim { it <= ' ' })
        return matcher.find()
    }

    fun findDate(value: String): String? {
        val originStrings = value.split("\n") as ArrayList<String>
        originStrings.reverse()
        originStrings.forEach { str ->
            if (str.isNotEmpty()) {
                val matcher: Matcher = datePattern.matcher(str.trim { it <= ' ' })
                if (matcher.find()) {
                    // первая найденная группа-дата, тут всё просто, она уже отформатирована
                    return matcher.group(1)
                }
            }
        }
        return null
    }
}