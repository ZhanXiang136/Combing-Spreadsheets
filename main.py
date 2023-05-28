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


def create_column_name(main_file):
    column_list = list(main_file.columns.values)
    while True:
        new_name = input('\nWhat do you want to name your new column?: ')
        if new_name not in column_list:
            return new_name

def main():
    print("Welcome to the Program")
    #look for spreadsheet
    main_file = 'main.xlsx'
    if not exists('main.xlsx'):
        main_file = check_file_existence('\nEnter the main spreadsheet: ')
    event_file = 'event.xlsx'
    if not exists('event.xlsx'):
        event_file = check_file_existence('\nEnter the event spreadsheet: ')

    #open up the spreadsheet
    df_main = determine_correct_pandas_conversion(main_file)
    df_event = determine_correct_pandas_conversion(event_file)


    df_main.columns = [x.lower() for x in df_main.columns]
    df_event.columns = [x.lower() for x in df_event.columns]
    for col in df_main.columns:
        if isinstance(df_main[col][1], str):
            df_main[col] = df_main[col].str.lower()
    for col in df_event.columns:
        if isinstance(df_event[col][1], str):
            df_event[col] = df_event[col].str.lower()

    invalid_list = []

    needed_column_names = ['First Name', 'Last Name']

    cal_what = input("(P)oints or (H)ours or (B)oth: ")
    if cal_what.lower().strip() == 'points' or cal_what.lower().strip() == 'p':
        look_for_points = True
        look_for_hours = False
        needed_column_names.append('points')
    elif cal_what.lower().strip() == 'hours' or cal_what.lower().strip() == 'h':
        look_for_points = False
        look_for_hours = True
        needed_column_names.append('hours')
    else:
        look_for_points = True
        look_for_hours = True
        needed_column_names.append('hours')
        needed_column_names.append('points')

    pattern = input("Sort by what? (E)mail, (O)sis?: ")
    if pattern.lower().strip() == 'osis' or pattern.lower().strip() == 'o':
        pattern = 'osis'
        needed_column_names.append('osis')
    elif pattern.lower().strip() == 'email' or pattern.lower().strip() == 'e':
        pattern = 'email'
        needed_column_names.append('email')

    needed_column_names = [x.lower() for x in needed_column_names]


    if all(item in list(df_event.columns.values) for item in needed_column_names) and all(
            item in list(df_main.columns.values) for item in needed_column_names):
        for index, row in df_event.iterrows():
            pattern_add = row[pattern]
            points_add = row['points'] if look_for_points else None
            hours_add = row['hours'] if look_for_hours else None
            first_name_add = row['first name']
            last_name_add = row['last name']

            found_row = df_main.loc[df_main[pattern] == pattern_add]
            len_of_found_rows = len(found_row)

            if len_of_found_rows >= 1:
                row_info_added = df_main.iloc[found_row.index[0]]
                first_name_added = row_info_added['first name']
                last_name_added = row_info_added['last name']
                if first_name_added.lower().strip() == first_name_add.lower().strip() and last_name_added.lower().strip() == last_name_add.lower().strip():
                    if look_for_points:
                        df_main.loc[df_main[pattern] == pattern_add, ['points']] = [points_add]
                    if look_for_hours:
                        df_main.loc[df_main[pattern] == pattern_add, ['hours']] = [hours_add]
                else:
                    print(f'\n{pattern} {pattern_add} has name {first_name_add} {last_name_add} in the event spreadsheet and {first_name_added} {last_name_added} in the main spreadsheet')
                    user_input = input("Press enter to add and continue, press x to skip and continue")
                    if user_input == 'x':
                        invalid_list.append(pattern_add)
                    else:
                        if look_for_points:
                            df_main.loc[df_main[pattern] == pattern_add, ['points']] = [points_add]
                        if look_for_hours:
                            df_main.loc[df_main[pattern] == pattern_add, ['hours']] = [hours_add]
            else:
                if len_of_found_rows == 0:
                    print(f'\n{pattern} with {pattern_add} not found in the main database')
                else:
                    print(f'\nFound multiple {pattern} with {pattern_add} in the main database')
                invalid_list.append(pattern_add)

        print(f'\nThe following {pattern} in the main spreadsheet are not added:')
        for i in invalid_list:
            print(i)

        df_main.to_excel('new.xlsx', index=False)

        print('\nProgram ended, thank you for using this code')

    else:
        print('\nYou are missing one or more the of the required column name')
        print(f'\nMain spreadsheet contains {list(df_main.columns.values)} and needed {needed_column_names}')
        print(f'\nEvent spreadsheet contains {list(df_event.columns.values)} and needed {needed_column_names}')


if __name__ == '__main__':
    main()
    print("Program Ended Done")

