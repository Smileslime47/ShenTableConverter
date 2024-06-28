import java.io.FileInputStream

fun main(args: Array<String>) {
    val path = "D:\\Projekt\\ShenTableConverter\\src\\test\\resources\\拼团排表(2).xlsx"
    PaiTableAnalyzer.analyze(path)
    ShenTableUpdater.update(path)
    val data = DataCache
    println(data)
}