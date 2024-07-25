package model

/**
 * 客户（吃谷人）定义
 */
class Customer {
    var nickname = ""
    var orderMap = HashMap<Goods, Int>()
    var isDepositPaid = true
    var isBalancePaid = false
}