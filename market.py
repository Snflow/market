import urllib
import simplejson
import xlwt
import sys
import getopt
from pprint import pprint

def get_price(typeID=34, scale='regionlimit', scaleID=10000043):
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


def get_history(typeID=34, regionID=10000043, days=10):
    api_address = "http://api.eve-marketdata.com/api/item_history2.json?char_name=market&region_ids="+str(regionID)+"&type_ids="+str(typeID)+"&days="+str(days)
    history_file = urllib.urlopen(api_address)
    history_json = history_file.read()
    history_file.close()
    history_data = simplejson.loads(history_json)

    total_volume = 0
    n_days = 0
    for single_day in history_data["emd"]["result"]:
        total_volume = total_volume + int(single_day["row"]["volume"])
        n_days = n_days + 1
    avg_volume = total_volume/max(1,n_days)
    if n_days == 0:
        avg_volume = 0
    return avg_volume


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


def main(argv):
    regionID = 10000043
    systemID = 30002187
    ID = 34
    volume_threshold = 100
    days = 10

    region_flag = False
    system_flag = False
    item_flag = False

    try:
        opts, args = getopt.getopt(argv,"r:v:s:d:i:", ["region=","volume=","system=","days=","item="])
    except getopt.GetoptError:
        print 'python market.py -r <regionID> -v <volume_threshold> -s <systemID> -d <days_for_volume> -i <item>'
        sys.exit(2)
    for opt,arg in opts:
        if opt in ("-r", "--region"):
            regionID = arg
            region_flag = True
        elif opt in ("-v", "--volume"):
            volume_threshold = int(arg)
        elif opt in ("-s", "--system"):
            systemID = arg
            system_flag = True
        elif opt in ("-d", "--days"):
            days = int(arg)
        elif opt in ("-i", "--item"):
            ID = arg
            item_flag = True
    if (system_flag == True and region_flag == False):
        print "Must specify the region ID which contains the system:", systemID
        exit()
    print "EVE market analyzer is generating the marketing date for:"
    print "    Region:", regionID
    if system_flag == True:
        print "    System:", systemID
        outfile = "system_"+str(systemID)+"&volume_"+str(volume_threshold)+"&days_"+str(days)+".xls"
    else:
        outfile = "region_"+str(regionID)+"&volume_"+str(volume_threshold)+"&days_"+str(days)+".xls"
    print "    The minimal average volume requirement in the past", days,"days is:", volume_threshold

    type_json = read_data()

    book = xlwt.Workbook(encoding="utf-8")
    sh = book.add_sheet("profit")
    sh.write(0,0,"Item")
    sh.write(0,1,"Type ID")
    sh.write(0,2,"Buy Price")
    sh.write(0,3,"Sell Price")
    sh.write(0,4,"Profit per Order")
    sh.write(0,5,"Average Volume")
    sh.write(0,6,"Total Profit Available")
    sh.write(0,7,"Profit Rate [%]")

    i = 0
    j = 1
    while type_json[i]["ID"] != 'end':
        ID = type_json[i]["ID"]
        name = type_json[i]["name"]

        if system_flag:
            (buy_price, sell_price) = get_price(typeID=ID, scale='usesystem', scaleID=systemID)
        else:
            (buy_price, sell_price) = get_price(typeID=ID, scaleID=regionID)
        if (buy_price != 0 and sell_price != 0):
            (profit, unit, profit_ratio) = unit_profit(buy_price, sell_price)
            avg_volume = get_history(typeID=ID, regionID=regionID, days=days)
            if avg_volume >= volume_threshold :
                profit_total = avg_volume * profit
                profit_out = str("{:8.2f}".format(profit))+unit
                profit_ratio_out = "{:8.2f}".format(profit_ratio)
                profit_total_out = str("{:8.2f}".format(profit_total))+unit

                sh.write(j,0,name)
                sh.write(j,1,ID)
                sh.write(j,2,buy_price)
                sh.write(j,3,sell_price)
                sh.write(j,4,profit_out)
                sh.write(j,5,avg_volume)
                sh.write(j,6,profit_total_out)
                sh.write(j,7,profit_ratio_out)
                print "Type ID:", ID, "|Item:", name, "|profit per order:", profit_out, "|profit ratio:", profit_ratio_out, "%", "|average volume:", avg_volume, "|total available profit:", profit_total_out
                j = j+1
        i = i+1

    book.save(outfile)

if __name__ == '__main__':
    main(sys.argv[1:])
