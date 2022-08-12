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
    column_list = list(file_to_be_added.columns.values)
    while True:
        new_name = input('\nWhat do you want to name your new column?: ')
        if new_name not in column_list:
            return new_name


def main():
    invalid_osis_list = []
    print("Please make sure that the file you're going to add have the following column: First Name, Last Name, Osis, Points")
    file_to_be_added = check_file_existence('\nEnter the file to be added: ')
    file_to_add = check_file_existence('\nEnter the file to add: ')

    df_to_be_added = determine_correct_pandas_conversion(file_to_be_added)
    df_to_add = determine_correct_pandas_conversion(file_to_add)

    new_column_name = create_column_name(df_to_be_added)
    df_to_be_added[new_column_name] = None

    if all(item in list(df_to_add.columns.values) for item in ['First Name', 'Last Name', 'Osis', 'Points']) and all(
            item in list(df_to_be_added.columns.values) for item in
            ['First Name', 'Last Name', 'Osis', 'Total Points']):
        index_of_total_point = list(df_to_be_added.columns.values).index('Total Points')
        for index, row in df_to_add.iterrows():
            osis_add = row['Osis']
            points_add = row['Points']
            first_name_add = row['First Name']
            last_name_add = row['Last Name']

            found_osis = df_to_be_added.loc[df_to_be_added['Osis'] == osis_add]
            length_of_osis_found = len(found_osis)

            if length_of_osis_found == 1:
                row_info_added = df_to_be_added.iloc[found_osis.index[0]]
                first_name_added = row_info_added['First Name']
                last_name_added = row_info_added['Last Name']
                total_point = row_info_added['Total Points'] if row_info_added['Total Points'] is None else 0
                if first_name_added.lower() == first_name_add.lower() and last_name_added.lower() == last_name_add.lower():
                    df_to_be_added.loc[df_to_be_added['Osis'] == osis_add, [new_column_name]] = [points_add]
                else:
                    print(f'\nOsis {osis_add} has name {first_name_add} {last_name_add} in the adding spreadsheet and {first_name_added} {last_name_added} in the added spreadsheet')
                    user_input = input("Press enter to add and continue, press x to skip and enter")
                    if user_input == 'x':
                        invalid_osis_list.append(osis_add)
                    else:
                        df_to_be_added.loc[df_to_be_added['Osis'] == osis_add, [new_column_name]] = [points_add]
            else:
                if length_of_osis_found == 0:
                    print(f'\nOSIS with {osis_add} not found in the main database')
                else:
                    print(f'\nFound multiple OSIS with {osis_add} in the main database')
                invalid_osis_list.append(osis_add)

        print('\nThe following osis in the to add spreadsheet are not added:')
        for i in invalid_osis_list:
            print(i)

        df_to_be_added['Total Points'] = df_to_be_added.iloc[:, index_of_total_point+1:].sum(axis=1)
        df_to_be_added.to_excel('new.xlsx', index=False)

        new_df = pd.read_excel('new.xlsx', sheet_name=0)
        print('\nProgram ended, thank you for using this code')

    else:
        print('\nYou are missing one or more the of the required column name')


if __name__ == '__main__':
    main()
