package Model

class Customer{
    var cn = ""
    var orderMap = HashMap<Goods,Int>()
    var isDepositPaid = false
    var isBalancePaid = false
}