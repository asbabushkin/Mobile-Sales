from openpyxl import load_workbook
from datetime import date
import pyodbc

server = 'DESKTOP-8QOVB79\SQLEXPRESS'  # имя сервера берем из бд (соединить...)
username = 'DESKTOP-8QOVB79\asbab'  # имя пользователя берем из бд
database = 'MobileSales'  # название подключаемой бд
conn = pyodbc.connect(DRIVER="{ODBC Driver 17 for SQL Server}", SERVER=server, DATABASE=database, UID=username,
                      Trusted_Connection="yes", autocommit=True)
cursor = conn.cursor()  # интерфейс для обращения к бд


def load_file():
    global wb
    global wb_sheets
    global file_name
    file_name = input('Введите имя файла ')
    wb = load_workbook(file_name)
    wb_sheets = wb.worksheets
    return wb, wb_sheets


# Преобразование текстового наименования оператора в числовой идентификатор
def get_operator_id(operator_name):
    temp = 'SELECT MobileOperatorID FROM MobileOperators WHERE MobileOperatorName = ' + '\'' + str(operator_name) + '\''
    cursor.execute(temp)
    return cursor.fetchone()[0]


# Преобразование текстового наименования тарифного плана в числовой идентификатор
def convert_rate_plan(rate_plan_name):
    temp = 'SELECT RatePlanID FROM RatePlans WHERE RatePlanName = ' + '\'' + str(rate_plan_name) + '\'' + ''
    cursor.execute(temp)
    return cursor.fetchone()[0]


def get_shop_id(shop_name_excel):
    temp = 'SELECT ShopID FROM Shops WHERE NameInExcel = ' + '\'' + str(shop_name_excel) + '\''
    cursor.execute(temp)
    print('get_shop_id:')
    # print(cursor.fetchall()[0][0])
    a = cursor.fetchone()[0]
    # for i in a:
    #     print(i)

    if a is None:
        return 999
    else:
        return a


# доделать проверку на наличие серийника в бд.
# В эксель при переносе данных в бд указывать не "да", а дату переноса

def load_sim_obtaining():
    print('Внесение прихода сим-карт в БД.')
    load_file()
    for sheet in wb_sheets:
        print(sheet)
        counter = 2
        while sheet['A' + str(counter)].value != None:
            if sheet['H' + str(counter)].value != 'Да':
                icc = sheet['A' + str(counter)].value
                icc = icc.replace(' ', '')
                operator_id = get_operator_id(sheet['C' + str(counter)].value)
                rate_plan_id = convert_rate_plan(sheet['D' + str(counter)].value)
                cost = sheet['E' + str(counter)].value
                receive_date = sheet['F' + str(counter)].value
                extra_mark = sheet['G' + str(counter)].value
                a = 'INSERT INTO SimProvision(ICC, MSISDN, MobileOperatorID, RatePlanID, Cost, ReceiveDate, ExtraMark) VALUES(' + '\'' + str(
                    icc) + '\'' + ',' + '\'' + '\'' + ',' + '\'' + str(
                    operator_id) + '\'' + ',' + '\'' + str(rate_plan_id) + '\'' + ',' + '\'' + str(
                    cost) + '\'' + ',' + '\'' + str(receive_date) + '\'' + ',' + '\'' + str(extra_mark) + '\'' + ')'
                cursor.execute(a)
                sheet['H' + str(counter)].value = 'Да'
                counter += 1
            else:
                print(sheet['A' + str(counter)].value + 'уже содержится в БД')
                counter += 1
    wb.save(file_name)
    return None


def load_sim_sales():
    pass


def main():
    menu = 'Выберите действие: \n 1 - загрузить приход сим-карт в БД \n 2 - Загрузить отгрузку сим-карт в БД \n 0 - выход '
    answer = int(input(menu))

    while answer != 0:
        if answer == 1:
            load_sim_obtaining()
            # print('Внесение прихода сим-карт в БД.')
            # load_file()
            # for sheet in wb_sheets:
            #     print(sheet)
            #     counter = 2
            #     while sheet['A' + str(counter)].value != None:
            #         if sheet['H' + str(counter)].value != 'Да':
            #             icc = sheet['A' + str(counter)].value
            #             operator_id = get_operator_id(sheet['C' + str(counter)].value)
            #             rate_plan_id = convert_rate_plan(sheet['D' + str(counter)].value)
            #             cost = sheet['E' + str(counter)].value
            #             receive_date = sheet['F' + str(counter)].value
            #             extra_mark = sheet['G' + str(counter)].value
            #             a = 'INSERT INTO SimProvision(ICC, MSISDN, MobileOperatorID, RatePlanID, Cost, ReceiveDate, ExtraMark) VALUES(' + '\'' + str(
            #                 icc) + '\'' + ',' + '\'' + '\'' + ',' + '\'' + str(
            #                 operator_id) + '\'' + ',' + '\'' + str(rate_plan_id) + '\'' + ',' + '\'' + str(
            #                 cost) + '\'' + ',' + '\'' + str(receive_date) + '\'' + ',' + '\'' + str(extra_mark) + '\'' + ')'
            #             cursor.execute(a)
            #             sheet['H' + str(counter)].value = 'Да'
            #             counter += 1
            #         else:
            #             print(sheet['A' + str(counter)].value + 'уже содержится в БД')
            #             counter += 1
            # wb.save(file_name)
            answer = int(input(menu))
        elif answer == 2:
            print('Загрузка продаж сим в БД ')
            load_file()
            for sheet in wb_sheets:
                counter = 2
                while sheet['D' + str(counter)].value != None:
                    icc = sheet['D' + str(counter)].value
                    print(icc)
                    shop_name_excel = sheet['F' + str(counter)].value
                    print(shop_name_excel)
                    transfer_date = sheet['G' + str(counter)].value
                    print(transfer_date)

                    shop_id = get_shop_id(shop_name_excel)
                    temp = 'INSERT INTO SimSales(ICC, ShopID, TransferDate) VALUES(' + '\'' + str(
                        icc) + '\'' + ', ' + '\'' + str(shop_id) + '\'' + ', ' + '\'' + str(transfer_date) + '\'' + ')'
                    print(temp)
                    cursor.execute(temp)
                    counter += 1
            answer = int(input(menu))


if __name__ == '__main__':
    main()
