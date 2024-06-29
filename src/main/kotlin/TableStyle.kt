import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class TableStyle(xssfWorkbook:XSSFWorkbook) {
    val borderStyle: CellStyle
    val currencyStyle: CellStyle
    val fieldStyle: CellStyle

    init {
        //borderStyle作为基础样式，以标题样式为基准修改
        borderStyle = xssfWorkbook.createCellStyle()
        borderStyle.cloneStyleFrom(DataCache.titleStyle)
        borderStyle.borderBottom = BorderStyle.THIN
        borderStyle.borderTop = BorderStyle.THIN
        borderStyle.borderLeft = BorderStyle.THIN
        borderStyle.borderRight = BorderStyle.THIN

        currencyStyle = xssfWorkbook.createCellStyle()
        currencyStyle.cloneStyleFrom(borderStyle)
        currencyStyle.dataFormat = 8

        fieldStyle = xssfWorkbook.createCellStyle()
        fieldStyle.cloneStyleFrom(borderStyle)
        fieldStyle.fillForegroundColor = IndexedColors.LIGHT_TURQUOISE.index
        fieldStyle.fillPattern = FillPatternType.SOLID_FOREGROUND;
    }
}