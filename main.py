import sys
import sqlite3
import re
from datetime import date
from PyQt6.QtSql import *
from PyQt6.QtCore import *
from PyQt6 import uic
from PyQt6.QtWidgets import *
Form_main, Window_main = uic.loadUiType("main_form.ui")
Form_add, Window_add = uic.loadUiType("add_form.ui")
Form_edit, Window_edit = uic.loadUiType("edit_form.ui")
Form_show, Window_show = uic.loadUiType("show_form.ui")
Form_confirm_deleting, Window_confirm_deleting = uic.loadUiType("confirm_deleting_form.ui")
Form_edit_row, Window_edit_row = uic.loadUiType("edit_row_form.ui")
Form_include_eg, Window_include_eg=uic.loadUiType("include_eg_form.ui")

database_name = "bd_var5.db"

def connect_db(database_name):
    con = QSqlDatabase.addDatabase("QSQLITE")
    con.setDatabaseName(database_name)
    if not con.open():
        print("Database Error: %s" % con.lastError().databaseText())
        sys.exit(1)
    print("Connection succeeded")
    con.close()

#---Отображение таблиц

def clicked_expert():
    table_model = QSqlTableModel()
    table_model.setTable("Expert_final")
    table_model.select()

    # loading all data
    while table_model.canFetchMore():
        table_model.fetchMore()
    table_model.rowCount()
    form_show.databaseTableView.setSortingEnabled(True)
    form_show.databaseTableView.setModel(table_model)
    form_show.databaseTableView.resizeColumnsToContents()
    form_show.databaseTableView.verticalHeader().setVisible(False)
    form_show.databaseTableView.hideColumn(6)
    form_show.databaseTableView.hideColumn(7)
    form_show.databaseTableView.hideColumn(9)
    form_show.databaseTableView.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

def clicked_region():
    table_model = QSqlTableModel()
    table_model.setTable("Reg_obl_city")
    table_model.select()
    # loading all data
    while table_model.canFetchMore():
        table_model.fetchMore()
    table_model.rowCount()
    #
    form_show.databaseTableView.setSortingEnabled(True)
    form_show.databaseTableView.setModel(table_model)
    form_show.databaseTableView.resizeColumnsToContents()
    form_show.databaseTableView.verticalHeader().setVisible(False)
    form_show.databaseTableView.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

def clicked_grnti():
    table_model = QSqlTableModel()
    table_model.setTable("grntirub")
    table_model.select()
    # loading all data
    while table_model.canFetchMore():
        table_model.fetchMore()
    table_model.rowCount()
    #
    form_show.databaseTableView.setSortingEnabled(True)
    form_show.databaseTableView.setModel(table_model)
    form_show.databaseTableView.resizeColumnsToContents()
    form_show.databaseTableView.verticalHeader().setVisible(False)
    form_show.databaseTableView.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)


#---Добавление данных в таблицу

def get_reg_data(database_name):
    con = sqlite3.connect(database_name)
    cur = con.cursor()
    data=[]
    data = cur.execute("SELECT DISTINCT region FROM Reg_obl_city").fetchall()
    cur.close()
    con.close()
    return data

def populate_region_combobox():
    list=[]
    data=get_reg_data(database_name)
    for x in data:
        list.append(str(x)[2:-3])
    form_add.regionComboBox.addItems(sorted(list))

def get_obl_data(database_name):
    con = sqlite3.connect(database_name)
    cur = con.cursor()
    data = []
    data = cur.execute("SELECT DISTINCT oblname FROM Reg_obl_city").fetchall()
    cur.close()
    con.close()
    return data

def get_city_data(database_name):
    con = sqlite3.connect(database_name)
    cur = con.cursor()
    data = []
    data = cur.execute("SELECT DISTINCT city FROM Reg_obl_city").fetchall()
    cur.close()
    con.close()
    return data

def populate_city_combobox():
    list=[]
    data=get_city_data(database_name)
    for x in data:
        list.append(str(x)[2:-3])
    form_add.cityComboBox.addItems(sorted(list))

def id_count():
    con = sqlite3.connect(database_name)
    cur = con.cursor()
    last_id = cur.execute("SELECT kod FROM Expert_final ORDER BY kod DESC LIMIT 1").fetchall()
    cur.close()
    con.close()
    return sum(last_id[0])

#---Проверки ввода

def check_name_input(name):
    name=" ".join(name.split())
    name=''.join(i for i in name if not i.isdigit())
    spaces = [i+1 for i,j in enumerate(name) if j==' ']
    capital_letters = [i for i,j in enumerate(name) if j.isupper()]
    if len(name)==0:
        print("Incorrect name")
        return False
    elif (name[-1]=='.'and name[-3]=='.' and name[-2].isupper() and name[-4].isupper() and name[0].isupper()):
        name = ''.join(name.split())
        name=name[:-4]+' '+name[-4:]
        return(name)
    elif (name[-1]=='.' and name[-2].isupper() and name[0].isupper()):
        name = ''.join(name.split())
        name = name[:-2] + ' ' + name[-2:]
        return(name)
    elif (len(name.split())==3 and len([idx for idx in range(len(name)) if name[idx].isupper()])==3 and spaces == capital_letters[1:3]):
        return (name)
    else:
        print("Incorrect name")
        return False

def check_grnti_input(grnti):
    grnti = re.sub('[^0-9.]', '', grnti)
    if len(grnti)==0:
        print("Incorrect grnti")
        return False
    elif len(grnti)==8:
        grnti1 = grnti[:8]
        grnti2 = ''
        grnti_search = grnti1[:2]
        return [grnti1,grnti2,grnti_search]
    elif len(grnti)==16:
        grnti1=grnti[:8]
        grnti2=grnti[8:]
        grnti_search = grnti1[:2]+" "+grnti2[:2]
        return [grnti1, grnti2, grnti_search]
    else:
        print("Incorrect grnti")
        return False

def same_person_check(name, city):
    con = sqlite3.connect(database_name)
    cur = con.cursor()
    res = cur.execute("""SELECT COUNT(*) FROM Expert_final WHERE name = '{}' AND city = '{}'""".format(name,city)).fetchall()
    res = sum(res[0])
    cur.close()
    con.close()
    return res

def get_city_region_dict():
    con = con = sqlite3.connect(database_name)
    cur = con.cursor()
    data_tuple = cur.execute("SELECT region,city FROM Expert_final ").fetchall()
    cur.close()
    con.close()
    data_list = list(set(data_tuple))
    region_list = []
    city_list = []
    for i in range(len(data_list)):
        region_list.append(data_list[i][0])
        city_list.append(data_list[i][1])
    res_dict = dict(zip(city_list,region_list))
    return res_dict

def region_city_check(region,city):
    res_dict = get_city_region_dict()
    return city in res_dict and region == res_dict[city]

#---Ввод данных

def get_input_data():
    name=str(form_add.nameLineAdd.text()).strip()
    region=str(form_add.regionComboBox.currentText()).strip()
    city=str(form_add.cityComboBox.currentText()).strip()
    grnti=str(form_add.grntiLineAdd.text()).strip()
    grnti_list=check_grnti_input(grnti)
    name = check_name_input(name)
    #?
    kod=id_count()+1
    today = date.today()
    input_date=today.strftime("%d.%m.%Y")
    if (grnti_list and name and region_city_check(region,city)):
        values = [kod, name, region, city, grnti_list[0], grnti_list[1], None, 0, input_date, grnti_list[2]]
        if same_person_check(name, city):
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Question)
            msg.setWindowTitle("Внимание")
            msg.setText("Человек с таким же ФИО в указанном городе уже находится в базе данных. Вы действительно хотите продолжить?")
            msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            buttonY = msg.button(QMessageBox.StandardButton.Yes)
            buttonY.setText('Да')
            buttonN = msg.button(QMessageBox.StandardButton.No)
            buttonN.setText('Нет')
            msg.exec()
            if msg.clickedButton() == buttonY:
                insert_into_db(values)
                form_add.nameLineAdd.clear()
                form_add.regionComboBox.clear()
                form_add.cityComboBox.clear()
                form_add.grntiLineAdd.clear()
                window_add.close()
                window_main.show()
            elif msg.clickedButton() == buttonN:
                msg.close()
        # msg.setInformativeText('More information')
        else:
            insert_into_db(values)
            form_add.nameLineAdd.clear()
            form_add.regionComboBox.clear()
            form_add.cityComboBox.clear()
            form_add.grntiLineAdd.clear()
            window_add.close()
            window_main.show()
    else:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setWindowTitle("Ошибка")
        msg.setText("Неверный ввод данных")
        #msg.setInformativeText('More information')
        msg.exec()

def insert_into_db(values):
    con = sqlite3.connect(database_name)
    cur = con.cursor()
    cur.execute("""INSERT INTO Expert_final VALUES(?,?,?,?,?,?,?,?,?,?)""", values).fetchall()
    con.commit()
    cur.close()
    con.close()

#---Работа с таблицами

def load_all_data():
    while table_model.canFetchMore():
        table_model.fetchMore()
    table_model.rowCount()

def get_selected_kod():
    rows_selected_kod = []
    indexes = form_edit.databaseEditTableView.selectionModel().selectedRows()
    for index in indexes:
        row = index.row()
        selected_rows = [proxy_model_input_date.index(row, col).data()
                            for col in range(proxy_model_input_date.columnCount())]
        rows_selected_kod.append(selected_rows[0])

    print(rows_selected_kod)
    indexes.clear()
    #selected_rows.clear()
    return rows_selected_kod

def get_selected():
    indexes = form_edit.databaseEditTableView.selectionModel().selectedRows()
    for index in indexes:
        row = index.row()
        selected_rows = [proxy_model_input_date.index(row, col).data()
                            for col in range(proxy_model_input_date.columnCount())]
    return selected_rows

def populate_edit_form():
    data = get_selected()
    form_edit_row.kodRowEdit.setText(str(data[0]))
    form_edit_row.kodRowEdit.setEnabled(False)
    form_edit_row.nameRowEdit.setText(str(data[1]))
    form_edit_row.regionRowEdit.setText(str(data[2]))
    form_edit_row.cityRowEdit.setText(str(data[3]))
    form_edit_row.grntiRowEdit.setText(str(str(data[4])+" "+str(data[5])))
    form_edit_row.inputDateRowEdit.setText(str(data[8]))

#def clicked_edit_expert():


def edit_db_row(values):
    query = QSqlQuery("""UPDATE Expert_final 
                            SET name = '{}',
                                region = '{}',
                                city = '{}',
                                grnti1 = '{}',
                                grnti2 = '{}',
                                key_words = '{}',
                                take_part = '{}',
                                input_date = '{}',
                                grnti_search = '{}' 
                            WHERE kod = {}""".format(values[1],values[2],values[3],values[4],values[5],values[6],values[7],values[8],values[9],values[0]))
    query.exec()


def edit_row():
    kod = int(form_edit_row.kodRowEdit.text())
    name = str(form_edit_row.nameRowEdit.text()).strip()
    name = check_name_input(name)
    region = str(form_edit_row.regionRowEdit.text()).strip()
    city = str(form_edit_row.cityRowEdit.text()).strip()
    grnti=str(form_edit_row.grntiRowEdit.text()).strip()
    grnti_list=check_grnti_input(grnti)
    input_date = str(form_edit_row.inputDateRowEdit.text()).strip()
    if (grnti_list and name and region_city_check(region,city)):
        values = [kod, name, region, city, grnti_list[0], grnti_list[1], None, 0, input_date, grnti_list[2]]
        print(values)
        edit_db_row(values)
        window_edit_row.close()
        table_model.select()
        load_all_data()
        window_edit.show()
    else:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setWindowTitle("Ошибка")
        msg.setText("Неверный ввод данных")
        # msg.setInformativeText('More information')
        msg.exec()

def delete_selected():
    rows_selected_kod=get_selected_kod()
    for kod in rows_selected_kod:
        query = QSqlQuery()
        query.prepare("DELETE FROM Expert_final WHERE kod=?")
        query.bindValue(0,kod)
        query.exec()
    table_model.select()
    load_all_data()

#---Экспертная группа

def include_in_eg():
    rows_selected_kod=get_selected_kod()
    for kod in rows_selected_kod:
        query = QSqlQuery()
        query.prepare("INSERT INTO Expert_group SELECT * FROM Expert_final WHERE kod=?")
        query.bindValue(0, kod)
        query.exec()
    delete_duplicates_in_eg()
    table_model_eg.select()
    form_include_eg.expertGroupTableView.resizeColumnsToContents()

def delete_duplicates_in_eg():
    query = QSqlQuery("DELETE FROM Expert_group WHERE rowid NOT IN (SELECT MIN(rowid) FROM Expert_group GROUP BY kod)")
    query.exec()

def delete_selected_eg():
    rows_selected_kod=[]
    indexes = form_include_eg.expertGroupTableView.selectionModel().selectedRows()
    for index in indexes:
        rows_selected_kod.append(index.data())
    for kod in rows_selected_kod:
        query = QSqlQuery()
        query.prepare("DELETE FROM Expert_group WHERE kod=?")
        query.bindValue(0, kod)
        query.exec()
    table_model_eg.select()

def confirm_eg():
    kod_list=[]
    con = sqlite3.connect(database_name)
    cur = con.cursor()
    res = cur.execute("SELECT kod FROM Expert_group").fetchall()
    con.commit()
    cur.close()
    con.close()
    for i in range(len(res)):
        kod_list.append(res[i][0])
    print(kod_list)
    if len(kod_list):
        con = sqlite3.connect(database_name)
        cur = con.cursor()
        for kod in kod_list:
            res = cur.execute("UPDATE Expert_final SET take_part = take_part + 1 WHERE kod = '{}'".format(kod)).fetchall()
            con.commit()
        res = cur.execute("DELETE FROM Expert_group")
        con.commit()
        cur.close()
        con.close()
        table_model_eg.select()
        table_model.select()
        load_all_data()
#---Переходы между окнами

def open_show_window():
    window_main.close()
    window_show.show()

def return_to_main_from_show():
    window_show.close()
    window_main.show()

def open_add_window():
    #window_main.close()
    window_add.show()

def return_to_main_from_add():
    form_add.nameLineAdd.clear()
    form_add.regionComboBox.clear()
    form_add.cityComboBox.clear()
    form_add.grntiLineAdd.clear()
    window_add.close()
    #window_main.show()

def open_edit_window():
    window_main.close()
    window_edit.show()

def return_to_main_from_edit():
    window_edit.close()
    window_main.show()

def open_confirm_window():
    #window_edit.close()
    window_confirm_deleting.show()

def return_to_edit_from_confirm():
    window_confirm_deleting.close()
    window_edit.show()

def open_edit_row_window():
    #window_edit.close()
    window_edit_row.show()

def return_to_edit_from_row():
    window_edit_row.close()
    window_edit.show()

def open_include_window():
    window_include_eg.show()

def return_to_edit_from_confirm_eg():
    window_include_eg.close()

app = QApplication([])
window_main = Window_main()
form_main = Form_main()
form_main.setupUi(window_main)
connect_db(database_name)
#form.databaseConnectButton.clicked.connect(lambda: connect_db(database_name))
#form.showDbButton.clicked.connect()
form_main.showDataButton.clicked.connect(open_show_window)

window_show = Window_show()
form_show = Form_show()
form_show.setupUi(window_show)

form_show.ExpertTableButton.clicked.connect(clicked_expert)
form_show.GrntiTableButton.clicked.connect(clicked_grnti)
form_show.RegionTableButton.clicked.connect(clicked_region)
form_show.returnToMainButton.clicked.connect(return_to_main_from_show)

form_main.addDataButton.clicked.connect(open_add_window)
form_main.editDataButton.clicked.connect(open_edit_window)
#form_main.editDataButton.clicked.connect(clicked_edit_expert)
form_main.addDataButton.clicked.connect(populate_region_combobox)
form_main.addDataButton.clicked.connect(populate_city_combobox)

window_add = Window_add()
form_add = Form_add()
form_add.setupUi(window_add)

form_add.addDataButton.clicked.connect(get_input_data)
form_add.returnToMainButton.clicked.connect(return_to_main_from_add)

window_edit = Window_edit()
form_edit = Form_edit()
form_edit.setupUi(window_edit)
window_confirm_deleting=Window_confirm_deleting()
form_confirm_deleting=Form_confirm_deleting()
form_confirm_deleting.setupUi(window_confirm_deleting)
table_model = QSqlTableModel()
table_model.setTable("Expert_final")
table_model.select()
load_all_data()
form_edit.databaseEditTableView.setSortingEnabled(True)
form_edit.databaseEditTableView.setModel(table_model)
form_edit.databaseEditTableView.resizeColumnsToContents()
form_edit.databaseEditTableView.verticalHeader().setVisible(False)
form_edit.databaseEditTableView.hideColumn(6)
#form_edit.databaseEditTableView.hideColumn(7)
form_edit.databaseEditTableView.hideColumn(9)
form_edit.databaseEditTableView.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
form_edit.databaseEditTableView.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)

source_model = table_model

proxy_model_kod = QSortFilterProxyModel()
proxy_model_kod.setSourceModel(source_model)
form_edit.databaseEditTableView.setModel(proxy_model_kod)
proxy_model_kod.setFilterKeyColumn(0)
proxy_model_kod.setFilterRegularExpression(QRegularExpression(form_edit.idFilterEdit.text()))
proxy_model_kod.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
form_edit.idFilterEdit.textChanged.connect(proxy_model_kod.setFilterRegularExpression)

proxy_model_name = QSortFilterProxyModel()
proxy_model_name.setSourceModel(proxy_model_kod)
form_edit.databaseEditTableView.setModel(proxy_model_name)
proxy_model_name.setFilterKeyColumn(1)
proxy_model_name.setFilterRegularExpression(QRegularExpression(form_edit.nameFilterEdit.text()))
proxy_model_name.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
form_edit.nameFilterEdit.textChanged.connect(proxy_model_name.setFilterRegularExpression)

proxy_model_region = QSortFilterProxyModel()
proxy_model_region.setSourceModel(proxy_model_name)
form_edit.databaseEditTableView.setModel(proxy_model_region)
proxy_model_region.setFilterKeyColumn(2)
proxy_model_region.setFilterRegularExpression(QRegularExpression(form_edit.regionFilterEdit.text()))
proxy_model_region.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
form_edit.regionFilterEdit.textChanged.connect(proxy_model_region.setFilterRegularExpression)

proxy_model_city = QSortFilterProxyModel()
proxy_model_city.setSourceModel(proxy_model_region)
form_edit.databaseEditTableView.setModel(proxy_model_city)
proxy_model_city.setFilterKeyColumn(3)
proxy_model_city.setFilterRegularExpression(QRegularExpression(form_edit.cityFilterEdit.text()))
proxy_model_city.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
form_edit.cityFilterEdit.textChanged.connect(proxy_model_city.setFilterRegularExpression)

proxy_model_grnti = QSortFilterProxyModel()
proxy_model_grnti.setSourceModel(proxy_model_city)
form_edit.databaseEditTableView.setModel(proxy_model_grnti)
proxy_model_grnti.setFilterKeyColumn(9)
proxy_model_grnti.setFilterRegularExpression(QRegularExpression(form_edit.grntiFilterEdit.text()))
proxy_model_grnti.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
form_edit.grntiFilterEdit.textChanged.connect(proxy_model_grnti.setFilterRegularExpression)

proxy_model_input_date = QSortFilterProxyModel()
proxy_model_input_date.setSourceModel(proxy_model_grnti)
form_edit.databaseEditTableView.setModel(proxy_model_input_date)
proxy_model_input_date.setFilterKeyColumn(8)
proxy_model_input_date.setFilterRegularExpression(QRegularExpression(form_edit.inputDateFilterEdit.text()))
proxy_model_input_date.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
form_edit.inputDateFilterEdit.textChanged.connect(proxy_model_input_date.setFilterRegularExpression)

form_confirm_deleting.confirmDeletingButton.clicked.connect(delete_selected)

window_edit_row=Window_edit_row()
form_edit_row=Form_edit_row()
form_edit_row.setupUi(window_edit_row)

window_include_eg=Window_include_eg()
form_include_eg=Form_include_eg()
form_include_eg.setupUi(window_include_eg)

form_edit.databaseEditTableView.doubleClicked.connect(populate_edit_form)
form_edit.databaseEditTableView.doubleClicked.connect(open_edit_row_window)
form_edit_row.saveChangesButton.clicked.connect(edit_row)
#form_edit_row.saveChangesButton.clicked.connect(return_to_edit_from_row)
form_edit_row.cancelButton.clicked.connect(return_to_edit_from_row)
form_edit.deleteDataButton.clicked.connect(open_confirm_window)
form_edit.returnToMainButton.clicked.connect(return_to_main_from_edit)
#form_edit.addExpertGroupButton.setEnabled(False)


form_confirm_deleting.confirmDeletingButton.clicked.connect(return_to_edit_from_confirm)
form_confirm_deleting.returnToEditButton.clicked.connect(return_to_edit_from_confirm)

form_include_eg.returnToEditButton.clicked.connect(return_to_edit_from_confirm_eg)
table_model_eg = QSqlTableModel()
table_model_eg.setTable("Expert_group")
table_model_eg.select()
form_include_eg.expertGroupTableView.setModel(table_model_eg)
form_include_eg.expertGroupTableView.setSortingEnabled(True)
form_include_eg.expertGroupTableView.sortByColumn(0,Qt.SortOrder.AscendingOrder)
form_include_eg.expertGroupTableView.resizeColumnsToContents()
form_include_eg.expertGroupTableView.verticalHeader().setVisible(False)
form_include_eg.expertGroupTableView.hideColumn(6)
form_include_eg.expertGroupTableView.hideColumn(9)
form_include_eg.expertGroupTableView.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
form_include_eg.expertGroupTableView.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
form_include_eg.removeExpertButton.clicked.connect(delete_selected_eg)
form_include_eg.confirmExpertGroupButton.clicked.connect(confirm_eg)

form_edit.confirmExpertGroupButton.clicked.connect(open_include_window)
form_edit.addExpertToGroupButton.clicked.connect(include_in_eg)



window_main.show()
app.exec()