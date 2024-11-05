import sqlite3
from sqlite3 import Error
from datetime import datetime
import win32com.client  
import os

class Tietokanta:
    def __init__(self, db_nimi='db_esim_1.db'):
        """Alustaa tietokannan ja luo yhteyden siihen."""
        self.conn = None
        self.cursor = None
        self.luo_yhteys(db_nimi)

    def luo_yhteys(self, db_nimi):
        """Luo yhteyden tietokantaan."""
        if os.path.exists(db_nimi):
            self.conn = sqlite3.connect(db_nimi)
            self.cursor = self.conn.cursor()
            self.luo_tietokanta()
        else:
            print(f"Tietokantaa {db_nimi} ei löydy.")

    def luo_tietokanta(self):
        """Luo päätaulun tietokannassa."""
        if self.conn:
            self.cursor.execute(''' 
            CREATE TABLE IF NOT EXISTS esimerkki (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                paivamaara DATETIME NOT NULL,
                lahettaja VARCHAR(255) NOT NULL,
                saaja VARCHAR(255) NOT NULL,
                teksti TEXT NOT NULL,
                aihe VARCHAR(255) NOT NULL
            )
            ''')
            self.conn.commit()

    def lisaa_tieto(self, lahettaja, saaja, teksti, aihe):
        """Lisää viestin tiedot tietokantaan."""
        if self.conn:
            paivamaara = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            self.cursor.execute('''
            INSERT INTO esimerkki (paivamaara, lahettaja, saaja, teksti, aihe)
            VALUES (?, ?, ?, ?, ?)
            ''', (paivamaara, lahettaja, saaja, teksti, aihe))
            self.conn.commit()

    def get_from_database(self):
        """Get data from the database's 'esimerkki' table."""
        if self.conn:
            self.cursor.execute("SELECT * FROM esimerkki")
            rows = self.cursor.fetchall()
            for row in rows:
                print(row)

    def remove_from_database(self, index):
        """Delete entry from database based on the given id-value."""
        if self.conn:
            self.cursor.execute("DELETE FROM esimerkki WHERE id = ?", (index,))
            self.conn.commit()
            print("Entry removed.")

    def sulje(self):
        """Sulje tietokannan yhteys."""
        if self.conn:
            self.conn.close()
            print("Connection closed.")


def laheta_sahkoposti(lahettaja, saaja, aihe, teksti):
    """Lähetä sähköposti Outlookin kautta."""
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0  # Sähköpostityyppi
    newmail = ol.CreateItem(olmailitem)  
    newmail.Subject = aihe  
    newmail.To = saaja
    newmail.Body = teksti
    newmail.Send() 

def main():
    """Main function. Contains all of the main code."""
    db = Tietokanta('esimerkki.db') 
    while True:
        print("\nMAIN MENU")
        print("1: Add entry")
        print("2: Remove entry")
        print("3: List database")
        print("4: Exit")
        user_input = input("Choice: ")

        if user_input == "1":
            lahettaja = input("Sender: ")
            saaja = input("Receiver: ")
            teksti = input("Message: ")
            aihe = input("Subject: ")
            db.lisaa_tieto(lahettaja, saaja, teksti, aihe)  
            laheta_sahkoposti(lahettaja, saaja, aihe, teksti) 
            print("Message added and sent.")
            
        elif user_input == "2":
            print("Database 'esimerkki' table:")
            db.get_from_database()  
            value_to_remove = int(input("Enter index (id) to remove: "))
            db.remove_from_database(value_to_remove) 
            
        elif user_input == "3":
            db.get_from_database() 

        elif user_input == "4":
            print("Quitting...")
            db.sulje() 
            break  

        else:
            print("Invalid choice.")

if __name__ == "__main__":
    main()
