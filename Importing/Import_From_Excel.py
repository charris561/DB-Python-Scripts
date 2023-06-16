'''Description: This program will be used to import data for inventory from 
an excel spreadsheet to a local MySQL server.
Author: Caleb Harris
Date: 5/11/2023
'''
#imports: 
import mysql.connector
from mysql.connector import Error
import ExcelFunctions
import CLI
import sys

#creates mysql server connection
def create_server_connection(host_name, user_name, user_password, db_name):
    connection = None
    try:
        connection = mysql.connector.connect(
            host=host_name,
            user=user_name,
            passwd=user_password,
            database=db_name
        )
        print("MySQL Database connection successful")
    except Error as err:
        print(f"Error: '{err}'")

    return connection

#executes passed in query
def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        result = cursor.fetchall()
        print("Query successful")
        return result
    except Error as err:
        print(f"Error: '{err}'")

#executes insert
def execute_insert(connection, query, data):
    cursor = connection.cursor()
    try:
        cursor.execute(query, data)
        connection.commit()
    except Error as err:
        print(f"Error: '{err}'")
        print(data)

#main function --> program entry point
def main():

    CLI.welcome_user()
    continue_program_execution = True

    while continue_program_execution:
        data = ExcelFunctions.getDataFromExcel()

        #database connection and querries
        pw = "PW Here!"
        connection = create_server_connection("localhost", "root", pw, "oit_services_inventory")
        
        #get data fields available for data mapping of excel sheet
        get_columns_query = "SHOW COLUMNS FROM assets"
        db_columns = execute_query(connection, get_columns_query)
        for i in range(len(db_columns)):
            db_columns[i] = db_columns[i][0]

        #perform data validation and data mapping. All rows of excel sheet should match a row of DB columns. 
        # Identifier for what the data is will the be column number in the 2D array
        mapped_data_list = []
        for i in range(len(data[0])):
            is_mappable = False
            for j in range(len(db_columns)):
                #format and compare strings by making them lowercase, replacing '#' with 'num' and 'expires' with date. Also removes '_' and ' '.
                if (db_columns[j].lower().replace('_', '') in data[0][i].lower().replace(' ', '').replace('#', 'num').replace('expires', 'date') and not is_mappable):
                    
                    #check for duplicate columns
                    if db_columns[j] in mapped_data_list:
                        if db_columns[j].lower().replace('_', '') == data[0][i].lower().replace(' ', '').replace('#', 'num').replace('expires', 'date') and not is_mappable:
                            is_mappable = True
                            mapped_data_list.append(db_columns[j])
                    else:
                        is_mappable = True
                        mapped_data_list.append(db_columns[j])
                        
            if (not is_mappable):
                raise Exception(f"Header: {data[0][i]} is not mappable.")
            
        #calculate location ID and floor code
        bldg_information = execute_query(connection, "SELECT * FROM BUILDINGS")
        location_id_index = mapped_data_list.index("location_id")
        building_index = mapped_data_list.index("building")
        area_num_index = mapped_data_list.index("area_num")
        floor_code_index = mapped_data_list.index("floor_code")
        for row in data[1:]:
            building = row[building_index]
            area_num = row[area_num_index]
            if building != None and area_num != None:

                #calculate building code
                building_code = ''
                for bldg in bldg_information:
                    if bldg[0] == building:
                        building_code = bldg[2]

                floor_code = building_code
                building_code += f"-{area_num}"
                row[location_id_index] = building_code

                #calculate floor code
                if f"{area_num}"[0] == "1":
                    floor_code += "-1st floor"
                elif f"{area_num}"[0] == "2":
                    floor_code += "-2nd floor"
                elif f"{area_num}"[0] == "3":
                    floor_code += "-3rd floor"
                
                row[floor_code_index] = floor_code


        #confirm all data wanted is present
        for db_column in db_columns:
            if not db_column in mapped_data_list:
                print(f"Database Column: {db_column} is not present in imported data. Would you like to proceed? ")
                answer = CLI.getUserConfirm()
                if (answer == False):
                    sys.exit()

        #get user confirmation for insertion
        print(f"About to insert {len(data) - 1} rows into DB. Would you like to proceed?")
        answer = CLI.getUserConfirm()
        if (answer == False):
            sys.exit()
        
        #build and execute insert query
        mapped_data_str = str(mapped_data_list).replace('[', '(').replace(']', ')').replace("'", '')
        insert_query = "INSERT INTO assets" + mapped_data_str + " VALUES (" + ("%s, "*len(mapped_data_list))[:-2] + ")"

        for i in range(1,len(data)):
            CLI.cmdProgressBar(i, len(data) - 1)
            execute_insert(connection, insert_query, data[i])
        print("\nOperation successfully completed!")

        print("Would you like to import more data? ")
        continue_program_execution = CLI.getUserConfirm()

#executes main only if program is run directly
if __name__ == "__main__":
    main()
