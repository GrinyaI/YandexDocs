from Excel import *
from Yandex import *


def main():
    DataBase_Name = "ОПД.xlsx"
    Group = "ПИН-221"
    Name = "Гриневич Илья"
    download_database(DATABASE_NAME=DataBase_Name)
    change_github(DATABASE_NAME=DataBase_Name, GROUP=Group, NAME=Name, NEW_LINK="https://github.com/")
    delete_database(DATABASE_NAME=DataBase_Name)
    upload_database(DATABASE_NAME=DataBase_Name)
    delete_file(DATABASE_NAME=DataBase_Name)


if __name__ == "__main__":
    main()
