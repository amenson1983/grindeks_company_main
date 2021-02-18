import psycopg2
from psycopg2 import OperationalError

def create_connection(db_name, db_user, db_password, db_host, db_port):
    connection = None
    try:
        connection = psycopg2.connect(
            database=db_name,
            user=db_user,
            password=db_password,
            host=db_host,
            port=db_port,
        )
        print("Connection to PostgreSQL DB successful")
    except OperationalError as e:
        print(f"The error '{e}' occurred")
    return connection

if __name__ == '__main__':
    db_name = 'medowl_grindex'
    db_user = 'grindex'
    db_password = 'xednirg'
    db_host = '62.149.15.123'
    db_port = '1433'
    quadra = create_connection(db_name, db_user, db_password, db_host, db_port)
