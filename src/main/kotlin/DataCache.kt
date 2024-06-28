import Model.Customer
import Model.Goods
import org.apache.poi.ss.usermodel.CellStyle

object DataCache {
    public var title = ""

    public var averageDeposit:Double = 0.0
    public var averageBalance:Double = 0.0
    public var paiTableEndLine = 0
    public val goodsMap = HashMap<String,Goods>()
    public val customerMap = HashMap<String,Customer>()
    public val customerList = ArrayList<Customer>()
    public lateinit var titleStyle: CellStyle

    object ColVal{
        public var roleCol:Int = 0
        public var priceFixCol:Int = 0
        public var customerStartCol:Int = 0
        public var customerEndCol:Int = 0
    }
}