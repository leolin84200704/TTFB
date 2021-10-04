import pandas as pd
from openpyxl import Workbook
import datetime


def main():
    fp = pd.read_csv('orderDetails.csv')
    date = datetime.date.today() - datetime.timedelta(days=2)
    print(date)
    filename = 'online-rate-leaderboard_' + str(date) + '_' + str(date) + '.csv'
    online_rate = pd.read_csv(filename)
    filename2 = 'efficiency-leaderboard_' + str(date) + '_' + str(date) + '.csv'
    efficiency = pd.read_csv(filename2)
    order_history = pd.read_csv('order.csv')
    # 計算上線率平均
    online_rate_filter = []
    su = 0
    for time in online_rate.上線率:
        if time > 0:
            online_rate_filter.append(time)
            su += time
    print("UE上線率: " + str(round(su/len(online_rate_filter), 2)) + "%")

    # 列出超過30分鐘門市

    time = '離線時間 (分鐘)'
    rate = '上線率'
    store = '商店'
    excel_file = Workbook()
    sheet = excel_file.active
    sheet['B2'] = 'Store'
    sheet['C2'] = 'time'
    subset = online_rate.loc[online_rate[time] > 30, (store, rate, time)]
    sub = subset.loc[subset[rate] > 0, (store, time)]
    i = 3
    for stores in sub[store]:
        sheet['A' + str(i)] = str(date.month) + '/' + str(date.day)
        sheet['B' + str(i)] = stores
        i += 1
    i = 3
    for times in sub[time]:
        sheet['C' + str(i)] = times
        i += 1
    # Ubereats 備餐超時明細
    time = "平均備餐時間"
    store = "商店"
    print('---------------------------')
    minute = sum(efficiency[time])/len(efficiency)/60
    print("UE平均備餐時間: " + str(round(minute, 2)))
    print('---------------------------')
    i += 2
    sheet['B' + str(i)] = '超時門市'
    sheet['C' + str(i)] = '超時單數'
    sheet['D' + str(i)] = '超時佔比'
    j = i + 1
    # 超過35分鐘的店家
    eff_sub = efficiency.loc[efficiency[time] > 2100, store]
    lis = []
    store = '餐廳'
    postpone = '已延長備餐時間？'
    for stores in eff_sub:
        sheet['A' + str(j)] = str(date.month) + '/' + str(date.day)
        sheet['B' + str(j)] = stores
        sub_order = order_history.loc[order_history[store] == stores, (store, postpone)]
        overtime = sum(sub_order[postpone])
        sheet['C' + str(j)] = overtime
        sheet['D' + str(j)] = '{:.0%}'.format(overtime/len(sub_order))
        lis.append(stores)
        j += 1
    j = i + 1
    excel_file.save(str(date) + '.xlsx')

    # FoodPanda Too Busy 店家

    reason = "Cancellation reason"
    store = ["Restaurant name"]
    print("FP拒單明細: ")
    print(fp[store][fp[reason] == "Too busy"])


if __name__ == '__main__':
    main()