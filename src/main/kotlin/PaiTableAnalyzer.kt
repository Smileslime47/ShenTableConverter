import Model.Customer
import Model.Goods
import PaiTableAnalyzer.Status.*
import org.apache.poi.ss.usermodel.CellType.*
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream

/**
 * 排表解析器
 */
object PaiTableAnalyzer {
    //字段扫描状态
    private val roleDataField = setOf("角色")
    private val priceFixField = setOf("调价")
    private var roleDataColGet = false
    private var priceFixColGet = false
    private var customerColGet = false

    //读取状态机
    enum class Status {
        NORMAL_SCAN,
        SCANNING_AVERAGE_DEPOSIT,
        SCANNING_AVERAGE_BALANCE,
        SCANNING_DATA_LINE,
    }

    fun analyze(path: String) {
        val inputStream = FileInputStream(path)
        val xssfWorkbook = XSSFWorkbook(inputStream)
        xssfWorkbook.use {
            val sheet: Sheet = it.getSheetAt(0)
            var readingStatus = NORMAL_SCAN
            val rowNum = sheet.lastRowNum
            var goodsNow:Goods = Goods()

            //行扫描
            rowLoop@ for (r in 0 until rowNum) {
                val row = sheet.getRow(r)
                //排表结束
                if(row == null) {
                    DataCache.paiTableEndLine = r
                    break@rowLoop
                }
                //排表起始
                if(r == 0) {
                    DataCache.title = row.first().stringCellValue
                    DataCache.titleStyle = row.first().cellStyle
                }

                val colNum = row.lastCellNum

                //每一行生成一个新的Goods对象
                if(readingStatus == SCANNING_DATA_LINE){
                    goodsNow = Goods()
                }

                //列扫描
                colLoop@ for (c in 0 until colNum) {
                    val cell = row.getCell(c) ?: break@colLoop

                    when (readingStatus) {
                        //字段扫描状态
                        NORMAL_SCAN -> {
                            when (cell.cellType!!) {
                                //数据字段
                                STRING -> {
                                    val text = cell.stringCellValue
                                    when {
                                        (text.equals("定金均价")) -> {
                                            readingStatus = SCANNING_AVERAGE_DEPOSIT
                                        }

                                        (text.equals("尾款均价")) -> {
                                            readingStatus = SCANNING_AVERAGE_BALANCE
                                        }

                                        (roleDataField.contains(text)) -> {
                                            DataCache.ColVal.roleCol = c
                                            roleDataColGet = true
                                        }

                                        (priceFixField.contains(text)) -> {
                                            DataCache.ColVal.priceFixCol = c
                                            priceFixColGet = true
                                        }
                                    }
                                }
                                //排表顺序
                                NUMERIC -> {
                                    val number = cell.numericCellValue.toInt()
                                    //第一配则说明开始记录排谷人
                                    if (number == 1) {
                                        DataCache.ColVal.customerStartCol = c
                                    }
                                }

                                //读到空则说明上一列单元格是配比结束
                                else -> {
                                    DataCache.ColVal.customerEndCol = c-1
                                    customerColGet = true
                                    break@colLoop
                                }
                            }
                        }

                        //读取定金均款
                        SCANNING_AVERAGE_DEPOSIT -> {
                            DataCache.averageDeposit = cell.numericCellValue
                            readingStatus = NORMAL_SCAN
                        }

                        //读取尾款均款
                        SCANNING_AVERAGE_BALANCE -> {
                            DataCache.averageBalance = cell.numericCellValue
                            readingStatus = NORMAL_SCAN
                        }

                        //读取数据行
                        SCANNING_DATA_LINE -> {
                            when{
                                //谷子名称
                                (c == DataCache.ColVal.roleCol) -> {
                                    goodsNow.roleName = cell.stringCellValue
                                    DataCache.goodsMap[goodsNow.roleName] = goodsNow
                                }

                                //谷子调价
                                (c == DataCache.ColVal.priceFixCol) -> {
                                    goodsNow.priceFix = cell.numericCellValue
                                }

                                //吃谷人
                                (c >= DataCache.ColVal.customerStartCol) -> {
                                    when (cell.cellType!!) {
                                        STRING,
                                        NUMERIC->{
                                            val cn:String = if(cell.cellType == STRING) {
                                                cell.stringCellValue
                                            }else {
                                                cell.numericCellValue.toString()
                                            }

                                            //cn为空则说明无后续
                                            if(cn.isBlank()){
                                                continue@colLoop
                                            }

                                            //表中不包括该吃谷人则添加
                                            if(!DataCache.customerMap.contains(cn)){
                                                val customer = Customer()
                                                customer.cn = cn
                                                DataCache.customerMap[cn] = customer
                                                DataCache.customerList.add(customer)
                                            }

                                            //谷子的吃谷人中不包括该吃谷人则添加
                                            val customer = DataCache.customerMap[cn]!!
                                            goodsNow.customers.add(customer)

                                            //吃谷数量初始化为0
                                            if(!customer.orderMap.contains(goodsNow)){
                                                customer.orderMap[goodsNow] = 0
                                            }

                                            //吃谷数量+1
                                            customer.orderMap[goodsNow] = customer.orderMap[goodsNow]!! + 1
                                        }

                                        else -> {
                                            break@colLoop
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //字段扫描完毕则开始读取数据
                if (roleDataColGet && priceFixColGet && customerColGet) {
                    readingStatus = SCANNING_DATA_LINE
                }
            }
        }
    }
}