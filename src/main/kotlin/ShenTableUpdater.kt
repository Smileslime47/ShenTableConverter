import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.ComparisonOperator
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellRangeAddressList
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream

/**
 * 肾表生成器
 */
class ShenTableUpdater(path: String) {
    private lateinit var tableStyle: TableStyle
    private lateinit var xssfWorkbook: XSSFWorkbook

    init {
        val inputStream = FileInputStream(path)
        xssfWorkbook = XSSFWorkbook(inputStream)
        tableStyle = TableStyle(xssfWorkbook)
    }

    fun update() {
        val outputStream = FileOutputStream("generated.xlsx")
        xssfWorkbook.use {
            val sheet = it.getSheetAt(0)
            var shenTableLine = DataCache.paiTableEndLine + 1

            shenTableLine = writeTitle(sheet, shenTableLine)

            shenTableLine = writeField(sheet, shenTableLine)

            readCustomer(sheet, shenTableLine)

            shenTableLine = writeCustomer(sheet, shenTableLine)

            shenTableLine = writeSummaryField(sheet, shenTableLine)

            shenTableLine = writeSummary(sheet, shenTableLine)

            for (i in 0..4) {
                sheet.autoSizeColumn(i)
            }

            it.write(outputStream)
        }
    }

    /**
     * 设置标题
     */
    private fun writeTitle(sheet: XSSFSheet, titleLineNum: Int): Int {
        val titleLine = sheet.getRow(titleLineNum)

        /**
         * 尝试合并标题单元格
         * 如果抛出异常说明可能已被合并，跳过该步骤
         */
        try {
            val titleRegion = CellRangeAddress(titleLine.rowNum, titleLine.rowNum, 0, 4)
            sheet.addMergedRegion(titleRegion)
        } catch (e: Exception) {
            //DO NOTHING
        }

        titleLine.heightInPoints = 50F

        for (i in 0..4) {
            val titleCell = titleLine.createCell(i)
            titleCell.setCellValue(
                if (DataCache.title.endsWith("排表"))
                    DataCache.title.substring(0, DataCache.title.length - 2) + "肾表"
                else
                    DataCache.title + "肾表"
            )
            titleCell.cellStyle = tableStyle.borderStyle as XSSFCellStyle
        }

        return titleLineNum + 1
    }

    /**
     * 设置字段行
     */
    private fun writeField(sheet: XSSFSheet, fieldLineNum: Int): Int {
        val fieldLine = sheet.getRow(fieldLineNum) ?: sheet.createRow(fieldLineNum)

        fieldLine.createCell(0).setCellValue("cn/xyid")
        fieldLine.createCell(1).setCellValue("定金")
        fieldLine.createCell(2).setCellValue("是否已肾")
        fieldLine.createCell(3).setCellValue("尾款")
        fieldLine.createCell(4).setCellValue("是否已肾")

        for (i in 0..4) {
            val cell = fieldLine.getCell(i) ?: fieldLine.createCell(i)
            cell.cellStyle = tableStyle.fieldStyle as XSSFCellStyle
        }

        return fieldLineNum + 1
    }

    private fun writeSummaryField(sheet: XSSFSheet, fieldLineNum: Int): Int {
        val fieldLine = sheet.getRow(fieldLineNum) ?: sheet.createRow(fieldLineNum)

        fieldLine.createCell(1).setCellValue("定金汇总")
        fieldLine.createCell(2).setCellValue("已肾定金")
        fieldLine.createCell(3).setCellValue("尾款汇总")
        fieldLine.createCell(4).setCellValue("已肾尾款")

        for (i in 0..4) {
            val cell = fieldLine.getCell(i) ?: fieldLine.createCell(i)
            cell.cellStyle = tableStyle.fieldStyle as XSSFCellStyle
        }

        return fieldLineNum + 1
    }

    /**
     * 读取已肾状态
     */
    private fun readCustomer(sheet: XSSFSheet, startLineNum: Int) {
        val maxRow = sheet.lastRowNum
        for (r in startLineNum until maxRow) {
            val row = sheet.getRow(r) ?: break

            val cnCell = row.getCell(0) ?: break
            val cn = cnCell.stringCellValue

            val depPaidCell = row.getCell(2) ?: break
            val depPaid = depPaidCell.booleanCellValue

            val balPaidCell = row.getCell(4) ?: break
            val balPaid = balPaidCell.booleanCellValue

            if (DataCache.customerMap.contains(cn)) {
                DataCache.customerMap[cn]!!.isDepositPaid = depPaid
                DataCache.customerMap[cn]!!.isBalancePaid = balPaid
            }
        }
    }

    private fun writeCustomer(sheet: XSSFSheet, startLineNum: Int): Int {
        /*
         * 设置数据有效性约束
         * 是否已肾一列应当仅从“已肾”和“未肾”两种情况中选择
         */
        val regionList1 = CellRangeAddressList(startLineNum, startLineNum + DataCache.customerList.size - 1, 2, 2)
        val regionList2 = CellRangeAddressList(startLineNum, startLineNum + DataCache.customerList.size - 1, 4, 4)
        val validationHelper = XSSFDataValidationHelper(sheet)
        val constraint = validationHelper.createExplicitListConstraint(arrayOf("已肾", "未肾"))
        val validation1 = validationHelper.createValidation(constraint, regionList1)
        val validation2 = validationHelper.createValidation(constraint, regionList2)

        sheet.addValidationData(validation1)
        sheet.addValidationData(validation2)

        /**
         * 设置格式规则
         * 已肾背景色应当自动为绿色，未肾背景色应当自动为红色
         */
        val region1 = CellRangeAddress(startLineNum, startLineNum + DataCache.customerList.size - 1, 2, 2)
        val region2 = CellRangeAddress(startLineNum, startLineNum + DataCache.customerList.size - 1, 4, 4)
        val formatting = sheet.sheetConditionalFormatting
        val rule1 = formatting.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"已肾\"")
        val rule2 = formatting.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"未肾\"")
        val pattern1 = rule1.createPatternFormatting()
        val pattern2 = rule2.createPatternFormatting()

        pattern1.fillBackgroundColor = IndexedColors.LIGHT_GREEN.index
        pattern1.fillPattern = FillPatternType.SOLID_FOREGROUND.code
        pattern2.fillBackgroundColor = IndexedColors.ROSE.index
        pattern2.fillPattern = FillPatternType.SOLID_FOREGROUND.code
        formatting.addConditionalFormatting(arrayOf(region1, region2), arrayOf(rule1, rule2))

        //写入吃谷数据
        for (r in 0 until DataCache.customerList.size) {
            val lineNum = startLineNum + r
            val row = sheet.getRow(lineNum) ?: sheet.createRow(lineNum)
            val customer = DataCache.customerList[r]

            for (c in 0..4) {
                val cell = row.getCell(c) ?: row.createCell(c)
                cell.cellStyle = tableStyle.borderStyle as XSSFCellStyle?
                when (c) {
                    0 -> {
                        cell.setCellValue(customer.nickname)
                    }

                    1 -> {
                        cell.cellType = CellType.NUMERIC
                        cell.cellStyle = tableStyle.currencyStyle as XSSFCellStyle?
                        cell.setCellValue(customer.orderMap.entries.fold(0.0) { acc, entry -> acc + DataCache.averageDeposit * entry.value })
                    }

                    2 -> {
                        cell.setCellValue(if (customer.isDepositPaid) "已肾" else "未肾")
                    }

                    3 -> {
                        cell.cellType = CellType.NUMERIC
                        cell.cellStyle = tableStyle.currencyStyle as XSSFCellStyle?
                        cell.setCellValue(customer.orderMap.entries.fold(0.0) { acc, entry ->
                            acc + (DataCache.averageBalance + entry.key.priceFix) * entry.value
                        })
                    }

                    4 -> {
                        cell.setCellValue(if (customer.isBalancePaid) "已肾" else "未肾")
                    }
                }
            }
        }

        return startLineNum + DataCache.customerList.size
    }

    private fun writeSummary(sheet: XSSFSheet, summaryLineNum: Int): Int {
        val summaryLine = sheet.getRow(summaryLineNum) ?: sheet.createRow(summaryLineNum)

        summaryLine.createCell(0)

        summaryLine.createCell(1).setCellValue(
            DataCache.customerList.fold(0.0) { acc1, customer ->
                acc1 + customer.orderMap.toList().fold(0.0) { acc2, pair ->
                    acc2 + DataCache.averageDeposit * pair.second
                }
            }
        )

        summaryLine.createCell(2).setCellValue(
            DataCache.customerList.fold(0.0) { acc1, customer ->
                acc1 + customer.orderMap.toList().fold(0.0) { acc2, pair ->
                    acc2 + if (customer.isDepositPaid) DataCache.averageDeposit * pair.second else 0.0
                }
            }
        )

        summaryLine.createCell(3).setCellValue(
            DataCache.customerList.fold(0.0) { acc1, customer ->
                acc1 + customer.orderMap.toList().fold(0.0) { acc2, pair ->
                    acc2 + (DataCache.averageBalance + pair.first.priceFix) * pair.second
                }
            }
        )

        summaryLine.createCell(4).setCellValue(
            DataCache.customerList.fold(0.0) { acc1, customer ->
                acc1 + customer.orderMap.toList().fold(0.0) { acc2, pair ->
                    acc2 + if (customer.isBalancePaid) (DataCache.averageBalance + pair.first.priceFix) * pair.second else 0.0
                }
            }
        )

        for (i in 0..4) {
            summaryLine.getCell(i).cellStyle = tableStyle.borderStyle as XSSFCellStyle
        }

        return summaryLineNum + 1
    }
}