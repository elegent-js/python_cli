import argparse
import os
import mysql.connector

def getArgs():
    """
    Get the command line arguments.
    """

    parser = argparse.ArgumentParser(description='Export database structure to Excel.')
    parser.add_argument('--host', required=True, help='Database host')
    parser.add_argument('--user', required=True, help='Database user')
    parser.add_argument('--password', required=True, help='Database password')
    parser.add_argument('--database', required=True, help='Database name')
    parser.add_argument('--port', type=int, default=3306, help='Database port')
    parser.add_argument('--output', default=os.path.join(os.getcwd(), 'generated_database_structure_to_excel.xlsx'), help='Path to save the generated Excel file (default: current working directory)')

    args = parser.parse_args()
    return args

def connect_to_database(host, user, password, database, port):
    """
    Establish a connection to the MySQL database.
    """
    return mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database,
        port=port
    )