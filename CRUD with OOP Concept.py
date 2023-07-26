import pyodbc
import csv

class ConnectPy:
    def __init__(self, connectionString: str) -> None:
        """Connects to Microsoft SQL Server"""
        self.server = 'DESKTOP-2I1END5\SQLEXPRESS'
        self.database = 'Student'
        self.connection_string = f"DRIVER={{SQL Server}};SERVER={self.server};DATABASE={self.database}; "
        self.connection = pyodbc.connect(self.connection_string)
        self.cursor = self.connection.cursor()

    def insert(self, contact_data):
        """Inserts values into the Contact  table and creates a csv file"""
        print('Insert')
        insert_query = 'INSERT INTO Contact (id, firstname, lastname, address) VALUES (?, ?, ?, ?)'
        self.cursor.execute(insert_query, contact_data)
        self.connection.commit()
        with open('contact.csv', mode='a', newline='') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(contact_data)

    def read(self):
        """Reads all the values from the table"""
        query = "SELECT * FROM Contact"
        self.cursor.execute(query)
        rows = self.cursor.fetchall()
        for row in rows:
            print(row)

    def update(self, id, firstname):
        """Updates the firstname having id 1 in the ocntact table"""
        print(f"Updating contact with id={id} to firstname={firstname}")
        update_query = "UPDATE Contact SET firstname = ? WHERE id = ?"
        try:
            self.cursor.execute(update_query, (firstname, id))
            self.connection.commit()
            print("Update successful")
            
        except pyodbc.Error as e:
            print(f"Error occurred during update: {str(e)}")

    def delete(self, id):
        """Deletes value from the contact table"""
        delete_query = 'DELETE FROM Contact WHERE id = ?'
        self.cursor.execute(delete_query, id)
        self.connection.commit()


    def close(self):
        self.cursor.close()
        self.connection.close()


if __name__ == "__main__":
    connection = ConnectPy("DatabaseConnected")
    contact_data = (4,"arrr","Koirala","Gwarko")
    connection.insert(contact_data)
    connection.read()
    connection.update( 3,"Pradeeptiiiiiii")
    connection.delete(3)
    connection.close()
