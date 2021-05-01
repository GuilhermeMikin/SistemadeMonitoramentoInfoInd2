import pyodbc

# try:
#     con_string = r'DRIVER = {Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=D:\Engenharia de Controle e Automação\IFIND 2\TRABALHO_T1\MeuProjeto\mypydatabase.mdb;'
    
#     conn = pyodbc.connect(con_string)
#     print("connected to database..")

# except pyodbc.Error as e:
#     print("Error in Connection", e.args)

# cursor = conn.cursor()
# cursor.execute('select * from customers')

# for row in cursor.fetchall():
#     print (row)

msa_drivers = [x for x in pyodbc.drivers()]
print(f'MS-Access Drivers: {msa_drivers}')

