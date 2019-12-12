import bs4
import pyodbc
import requests
import os
import openpyxl
wb = openpyxl.Workbook()
wb.create_sheet(index=0, title="Sales Data")
conn = pyodbc.connect("Driver={SQL Server};"
                      "Server=157.201.228.85;"
                      "Database=AdventureWorks2008R2;"
                      "UID=SQLstudent;"
                      "pwd=sqlbyu1daho;"
                      "Trusted_Connection=no;")

cursor = conn.cursor()
cursor.execute("SELECT p.Name, SUM(s.LineTotal)"
               "FROM AdventureWorks2008R2.Sales.SalesOrderDetail s"
               "INNER JOIN AdventureWorks2008R2.Production.Product p"
               "ON s.ProductID"
               "INNER JOIN AdventureWorks2008R2.Sales.SalesOrderHeader h"
               "ON s.SalesOrderID = h.SalesOrderID"
               "WHERE h.OrderDate BETWEEN '1/1/2008' AND '7/1/2008'"
               "GROUP BY p.Name"
               "ORDER BY SUM(s.LineTotal)DESC;")

for row in cursor:

    mydict = {rows[0]:rows[1] for rows in cursor}

    i = 1
    for k, v in mydict.items():
        sheet["A"+str(i)] = k
        sheet["B"+str(i)] = v
        i += 1

os.chdir("c:\\Users\\Skywalker\\Desktop\\Python")
wb.save("newdata.xlsx")


