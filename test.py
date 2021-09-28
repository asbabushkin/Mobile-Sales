import pyodbc

server = 'DESKTOP-8QOVB79\SQLEXPRESS'  # имя сервера берем из бд (соединить...)
username = 'DESKTOP-8QOVB79\asbab'  # имя пользователя берем из бд
database = 'MobileSales'  # название подключаемой бд
conn = pyodbc.connect(DRIVER="{ODBC Driver 17 for SQL Server}", SERVER=server, DATABASE=database, UID=username,
                      Trusted_Connection="yes", autocommit=True)
cursor = conn.cursor()  # интерфейс для обращения к бд

icc = '89701022405718028   '
shop_id = 99
transfer_date = '2021.09.15'
# shop_name = 'Попово ИП Емельянова Г.Е.'
# a = 'SELECT ShopID FROM Shops WHERE NameInExcel = ' + '\'' + shop_name + '\''
# cursor.execute(a)
# a = cursor.fetchone()[0]
a = 'INSERT INTO SimSales(ICC, ShopID, TransferDate) VALUES(' + '\'' + icc + '\'' + ', ' + '\'' + str(shop_id) + '\'' + ', ' + '\'' + transfer_date + '\'' + ')'
print(a)
