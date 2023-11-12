from Excel import *

def main():
    DataBase_Name = "ОПД.xlsx"
    Group = "ПИН-222"
    Name = "Шляхтин Роман"
    Id = "id123512"
    set_telegram_id(DATABASE_NAME=DataBase_Name, GROUP=Group, NAME=Name, NEW_TELEGRAM_ID=Id)


if __name__ == "__main__":
    main()
