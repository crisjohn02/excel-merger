# import mysql.connector
# from mysql.connector import Error
#
# try:
#     connection = mysql.connector.connect(host='localhost',
#                                          database='impress3',
#                                          user='root',
#                                          password='')
#
#     if connection.is_connected():
#         db_info = connection.get_server_info()
#         print("Connected to MySQL Server version ", db_info)
#         cursor = connection.cursor()
#         cursor.execute("select database();")
#         record = cursor.fetchone()
#         print("You're connected to database: ", record)
#
# except Error as e:
#     print("Error while connecting to MySQL", e)
# finally:
#     if connection.is_connected():
#         cursor.close()
#         connection.close()
#         print("MySQL connection is closed")

from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy import select
import models
import json

SQLALCHEMY_DATABASE_URI = "mysql+pymysql://root:@localhost:3306/impress3?charset=utf8mb4"

mysql_engine = create_engine(SQLALCHEMY_DATABASE_URI, isolation_level="AUTOCOMMIT")


with Session(mysql_engine) as session:
    stmt = select(models.Record).where(models.Record.item_id == "67e11b46-709a-4be8-941c-c96b55cea42a").limit(1)

    for record in session.scalars(stmt):
        print(record.data.get("main"))
    # c = [r for r in [x.item.primes for x in session.scalars(stmt)]]
    # print(c)
