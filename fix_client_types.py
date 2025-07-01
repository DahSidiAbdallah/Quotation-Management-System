import sqlite3

db = sqlite3.connect('clients.db')
c = db.cursor()
c.execute("UPDATE clients SET client_type='ciment' WHERE UPPER(name)='TASIAST'")
c.execute("UPDATE clients SET client_type='beton' WHERE UPPER(name)='TND'")
db.commit()
for row in c.execute('SELECT name, client_type FROM clients'):
    print(row)
db.close()
