import mysql.connector

class Database:
  def connect_to_database():
    conn = mysql.connector.connect(
      host= "localhost",
      user= "root",
      password= "",
      database= "absensi_tkj1"
    )
    return conn
  
  # digunakan untuk connect database