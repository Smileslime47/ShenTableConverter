import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream

/**
 * 肾表生成器
 */
object ShenTableUpdater {
    lateinit var borderStyle:CellStyle

    fun update(path: String) {
        val inputStream = FileInputStream(path)
        val outputStream = FileOutputStream("generated.xlsx")
        val xssfWorkbook = XSSFWorkbook(inputStream)
        xssfWorkbook.use {
            val sheet = it.getSheetAt(0)
            var shenTableLine = DataCache.paiTableEndLine + 1

            setBorderStyle(xssfWorkbook)

            setTitle(xssfWorkbook,shenTableLine++)

            setField(sheet.getRow(shenTableLine++))

            readCustomer(sheet,shenTableLine)

            writeCustomer(sheet,shenTableLine)

            it.write(outputStream)
        }
    }

    private fun setBorderStyle(xssfWorkbook:XSSFWorkbook){
        borderStyle = xssfWorkbook.createCellStyle()
        borderStyle.cloneStyleFrom(DataCache.titleStyle)
        borderStyle.borderBottom = BorderStyle.THIN
        borderStyle.borderTop = BorderStyle.THIN
        borderStyle.borderLeft = BorderStyle.THIN
        borderStyle.borderRight = BorderStyle.THIN
    }

    /**
     * 设置标题
     */
    private fun setTitle(xssfWorkbook:XSSFWorkbook, shenTableStartLine:Int){
        try{
            val titleRegion = CellRangeAddress(shenTableStartLine,shenTableStartLine,0,4)
            val sheet = xssfWorkbook.getSheetAt(0)
            sheet.addMergedRegion(titleRegion)

            val startRow = sheet.createRow(shenTableStartLine)

            for(i in 0..4){
                val titleCell = startRow.createCell(i)
                titleCell.setCellValue(DataCache.title+"肾表")
                titleCell.cellStyle= borderStyle as XSSFCellStyle
            }
        }catch (e:Exception){
            //DO NOTHING
        }
    }

    /**
     * 设置字段行
     */
    private fun setField(fieldLine: Row){
        fieldLine.createCell(0).setCellValue("cn/xyid")
        fieldLine.createCell(1).setCellValue("定金")
        fieldLine.createCell(2).setCellValue("是否已肾")
        fieldLine.createCell(3).setCellValue("尾款")
        fieldLine.createCell(4).setCellValue("是否已肾")

        for(i in 0..4){
            val cell = fieldLine.getCell(i)
            cell.cellStyle= borderStyle as XSSFCellStyle
        }
    }

    /**
     * 读取已肾状态
     */
    private fun readCustomer(xssfSheet: XSSFSheet,startLine:Int) {
        val maxRow = xssfSheet.lastRowNum
        for(r in startLine until maxRow){
            val row = xssfSheet.getRow(r) ?: break

            val cnCell = row.getCell(0)?:break
            val cn = cnCell.stringCellValue

            val depPaidCell = row.getCell(2)?:break
            val depPaid = depPaidCell.booleanCellValue

            val balPaidCell = row.getCell(4)?:break
            val balPaid = balPaidCell.booleanCellValue

            if(DataCache.customerMap.contains(cn)){
                DataCache.customerMap[cn]!!.isDepositPaid = depPaid
                DataCache.customerMap[cn]!!.isBalancePaid = balPaid
            }
        }
    }

    private fun writeCustomer(xssfSheet: XSSFSheet,startLine:Int) {
        for(r in 0 until DataCache.customerList.size){
            val lineNum = startLine + r
            val row = xssfSheet.getRow(lineNum) ?: xssfSheet.createRow(lineNum)
            val customer = DataCache.customerList[r]

            for(c in 0 .. 4){
                val cell = row.getCell(c)?:row.createCell(c)
                cell.cellStyle = borderStyle as XSSFCellStyle?
                when(c){
                    0 -> {
                        cell.setCellValue(customer.cn)
                    }
                    1 -> {
                        cell.setCellValue(customer.orderMap.entries.fold(0.0){
                            acc,entry -> acc + DataCache.averageDeposit * entry.value})
                    }
                    2 -> {
                        cell.setCellValue(customer.isDepositPaid)
                    }
                    3 -> {
                        cell.setCellValue(customer.orderMap.entries.fold(0.0){
                            acc,entry -> acc + (DataCache.averageBalance + entry.key.priceFix) * entry.value
                        })
                    }
                    4 -> {
                        cell.setCellValue(customer.isBalancePaid)
                    }
                }
            }
        }
    }
}