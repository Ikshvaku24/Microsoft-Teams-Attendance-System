# import pymysql
def removing(j):
    database_list = list(list(i) for i in j)
    database_list.pop(database_list.index(['information_schema']))
    database_list.pop(database_list.index(['mysql']))
    database_list.pop(database_list.index(['performance_schema']))
    database_list.pop(database_list.index(['sys']))
    return database_list

def database_view():
    try:
        db = pymysql.connect(host='localhost', user='iksh', passwd='iksh', autocommit=True)
        cur = db.cursor()
        if cur.connection:
            print('succesfully connected')
        cur.execute('show databases;')
        return removing(cur.fetchall())
    except Exception as e:
        print(e)
    finally:
        db.close()
def table_list(database):
    try:
        db = pymysql.connect(host='localhost', user='iksh', passwd='iksh', autocommit=True)
        cur = db.cursor()
        if cur.connection:
            print('succesfully connected')
        cur.execute(f'USE `{database}`;')
        cur.execute(f'show tables;')
        return list(list(i) for i in cur.fetchall())
    except Exception as e:
        print(e)
    finally:
        db.close()


def database_creation(classlist_file):
    try:
        db = pymysql.connect(host='localhost', user='iksh', passwd='iksh', autocommit=True)
        cur = db.cursor()
        if cur.connection:
            print('succesfully connected')
        cur.execute(f'CREATE DATABASE IF NOT EXISTS {classlist_file};')
        cur.execute('show databases;')

        print(removing(cur.fetchall()))
    except Exception as e:
        print(e)
    finally:
        db.close()


def name_of_database(classlist_file):
    l = [i for i in range(len(classlist_file)) if classlist_file[i] == '/']
    classlist_file = classlist_file[l.pop() + 1:classlist_file.find('.xlsx')]
    return classlist_file


def table_creation(table_name, time_use, classlist_file):
    try:
        db = pymysql.connect(host='localhost', user='iksh', passwd='iksh', autocommit=True)
        cur = db.cursor()
        if cur.connection:
            print('succesfully connected')
    except Exception as e:
        print(e)
    print(classlist_file)
    try:
        cur.execute(f'USE `{classlist_file}`;')
        cur.execute(f'CREATE TABLE IF NOT EXISTS `{table_name}` (\
        SNO int, Name varchar(20), time_spent TIME);')
        n = len(time_use)
        for i in range(n):
            cur.execute(f"INSERT INTO `{table_name}` VALUES ({i + 1}, '{time_use[i][0]}', '{time_use[i][1]}');")
    except Exception as e:
        print(e)
    finally:
        db.close()
