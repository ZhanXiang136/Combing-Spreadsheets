import pandas as pd
from os.path import exists

def check_file_existence(text):
    while True:
        user_input = input(f"{text}")
        if exists(user_input):
            return user_input
        else:
            print("Invalid file path")


def determine_correct_pandas_conversion(file):
    extension = file.split('.')[1]
    if extension == 'xlsx':
        return pd.read_excel(file, sheet_name=0)
    elif extension == 'csv':
        return pd.read_csv(file)
    elif extension == 'tsv':
        return pd.read_csv(file, delimiter='\t')
    else:
        print(f'\nError: Invalid extension {extension}')


def create_column_name(file_to_be_added):
    return 'event'
    column_list = list(file_to_be_added.columns.values)
    while True:
        new_name = input('\nWhat do you want to name your new column?: ')
        if new_name not in column_list:
            return new_name


def main():
    invalid_First_Name_list = []
    #print("Please make sure that the file you're going to add have the following column: First Name, Last Name, Osis, Points")
    #file_to_be_added = check_file_existence('\nEnter the file to be added: ')
    #file_to_add = check_file_existence('\nEnter the file to add: ')

    df_main = pd.read_excel('main.xlsx', sheet_name=0)#determine_correct_pandas_conversion(file_to_be_added)
    df_event = pd.read_excel('event.xlsx', sheet_name=0)#determine_correct_pandas_conversion(file_to_add)

    new_column_name = create_column_name(df_main)
    df_main[new_column_name] = None

    if all(item in list(df_event.columns.values) for item in ['First Name', 'Last Name', 'Points', 'Hours']) and all(
            item in list(df_main.columns.values) for item in
            ['First Name', 'Last Name', 'Osis', 'Total Points']):
        index_of_total_point = list(df_main.columns.values).index('Total Points')
        for index, row in df_event.iterrows():
            found_row = False
            points_add = row['Points']
            first_name_add = row['First Name'].capitalize()
            last_name_add = row['Last Name']
            hours_add = row['Hours']

            found_First_Name = df_main.loc[df_main['First Name'] == first_name_add]
            length_of_First_Name_found = len(found_First_Name)

            if length_of_First_Name_found >= 1:
                for i in range(length_of_First_Name_found):
                    row_info_added = df_main.iloc[found_First_Name.index[i]]
                    last_name_added = row_info_added['Last Name']
                    try:
                        if last_name_added.lower().strip() == last_name_add.lower().strip():
                            df_main.loc[(df_main['First Name'] == first_name_add) & (df_main['Last Name'] == last_name_add), [new_column_name]] = [points_add]
                            df_main.loc[(df_main['First Name'] == first_name_add) & (df_main['Last Name'] == last_name_add), ['Hours']] = [hours_add]
                            found_row = True
                            break
                    except Exception as e:
                        print(e)
                        break
                if not found_row:
                    invalid_First_Name_list.append(f'{first_name_add} {last_name_add} {points_add} + {hours_add}')

            else:
                if length_of_First_Name_found == 0:
                    print(f'\nFirst_Name with {first_name_add} not found in the main database')
                else:
                    print(f'\nFound multiple First_Name with {first_name_add} in the main database')
                invalid_First_Name_list.append(f'{first_name_add} {last_name_add} {points_add} + {hours_add}')

        print('\nThe following First_Name in the main spreadsheet are not added:')
        for i in invalid_First_Name_list:
            print(i)

        df_main['Total Points'] = df_main.iloc[:, index_of_total_point+1:].sum(axis=1)
        df_main.to_excel('new.xlsx', index=False)

        print('\nProgram ended, thank you for using this code')

    else:
        print('\nYou are missing one or more the of the required column name')

if __name__ == '__main__':
    main()
    print("Done")
