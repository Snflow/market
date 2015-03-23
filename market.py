import urllib
import simplejson
import xlwt
from pprint import pprint

def get_price(typeID=34, scale='usesystem', scaleID=30002187):
    # Generate api address on eve-central
    api_address = "http://api.eve-central.com/api/marketstat/json?typeid="+str(typeID)+"&"+scale+"="+str(scaleID)
    # Receive raw market JSON strings.
    market_file = urllib.urlopen(api_address)
    market_json = market_file.read()
    market_file.close()
    # Un-serialize the JSON data to a Python dict.
    market_data = simplejson.loads(market_json)
    # Get buy and sell prices.
    buy_price = market_data[0]["buy"]["max"]
    sell_price = market_data[0]["sell"]["min"]
    return(buy_price, sell_price)


def broker_tax(buy_price, sell_price):
    # Broker fee ratio, affected by both skill and standings.
    broker_ratio = 0.0075
    # Tax ratio, only affected by skill.
    tax_ratio = 0.0075
    # Broker fees for buy and sell.
    broker_buy = broker_ratio*buy_price
    broker_sell = broker_ratio*sell_price
    broker = broker_buy + broker_sell
    # Tax for sell.
    tax = tax_ratio*sell_price
    return(broker, tax)


def unit_profit(buy_price, sell_price):
    non_zero = 0.0000001
    (broker, tax) = broker_tax(buy_price, sell_price)

    profit = sell_price - buy_price - broker - tax
    profit_ratio = profit/(buy_price+non_zero)*100

    if profit/1000000 < 1.0:
        if profit/1000 < 1.0:
            profit = profit
            unit = " isk"
        else:
            profit = profit / 1000
            unit = "K isk"
    else:
        profit = profit / 1000000
        unit = "M isk"
    return (profit, unit, profit_ratio)


def read_data():
    file = open('data')
    type_json = simplejson.load(file)
    file.close()
    return type_json


def main():
    type_json = read_data()

    book = xlwt.Workbook(encoding="utf-8")
    sh = book.add_sheet("profit")
    sh.write(0,0,"Item")
    sh.write(0,1,"Type ID")
    sh.write(0,2,"Buy Price")
    sh.write(0,3,"Sell Price")
    sh.write(0,4,"Profit/order")
    sh.write(0,5,"Profit_rate")

    i = 0
    j = 1
    while type_json[i]["ID"] != 'end':
        ID = type_json[i]["ID"]
        name = type_json[i]["name"]

        (buy_price, sell_price) = get_price(typeID=ID, scale='regionlimit', scaleID=10000043)
        if (buy_price != 0 and sell_price != 0):
            (profit, unit, profit_ratio) = unit_profit(buy_price, sell_price)
            profit_out = str("{:8.2f}".format(profit))+unit
            profit_ratio_out = str("{:8.2f}".format(profit_ratio))+"%"
            sh.write(j,0,name)
            sh.write(j,1,ID)
            sh.write(j,2,buy_price)
            sh.write(j,3,sell_price)
            sh.write(j,4,profit_out)
            sh.write(j,5,profit_ratio_out)
            print "Type ID:", ID, "Item:", name, "profit per order:", profit_out, "profit ratio:", profit_ratio_out
            j = j+1
        i = i+1

    book.save("profit.xls")

if __name__ == '__main__':
    main()
