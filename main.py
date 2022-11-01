
import pandas as pd
import PySimpleGUI as ps


file_data = pd.read_excel('D:\\Me\\Python\\Excel\\Data.xlsx')
table_header = file_data.columns.tolist()
id_number = file_data['ID'].values
sanpham = file_data['Tên sản phẩm'].values


layout = [
    [ps.Text("Bảng điều khiển")],
    [ps.Text("ID",size =20), ps.Input(key="ID")],
    [ps.Text("Tên sản phẩm",size =20), ps.Input(key="Tên sản phẩm")],
    [ps.Text("Số lượng",size =20), ps.Input(key="Số lượng")],
    [ps.Text("Giá",size =20), ps.Input(key="Giá")],
    [ps.Text("Đã bán",size =20), ps.Input(key="Đã bán")],
    [ps.Button("Data", button_color='Blue'),
    ps.Button("Save", button_color='green'), 
    ps.Button("Search", button_color='yellow'), 
    ps.Button("Delete", button_color='red'), 
    ps.Button("Exit", button_color='black'),
    ps.Button("Reset", button_color='gray')],
    [ps.Table(values = " ",    
            headings = table_header,
            key="Table",
            row_height = 30,
            justification = 'center',
            expand_x= True,
            expand_y= True,)]
    ]


window = ps.Window("Thông tin sản phẩm",layout)

def clear_input():
    for key in values:
        window[key](" ")
    return None
    

while True:
    event,values = window.read()
    print(values)
    if event == ps.WIN_CLOSED or event == "Exit":
        break
    if event == "Save":
        check_id = int(values['ID'])
        check_sp = values['Tên sản phẩm']
        if check_id == "" or check_sp == "":
            ps.popup("Vui lòng nhập đầy đủ thông tin sản phẩm")
        elif check_id in id_number :
            indexa = file_data.loc[file_data['ID'] == check_id].index.to_list()[0]
            header_list = list(file_data.columns.values)
            for key in header_list : 
                file_data.loc[indexa,key] = values[key]
            file_data.to_excel('D:\\Me\\Python\\Excel\\Data.xlsx', index=False)
            ps.Popup("Sửa thành công")
            clear_input()
        else :
            del values['Table']
            file_data = file_data.append(values,ignore_index=True)
            file_data.to_excel('D:\\Me\\Python\\Excel\\Data.xlsx', index=False)
            ps.Popup("Thêm thành công")
            clear_input()
    if event == "Search":
        check_sp = values['Tên sản phẩm']
        check_id = values['ID']
        if check_id =='' and check_sp =='' :
            ps.popup('Vui lòng nhập thông tin sản phẩm ')
        else :
            if check_sp in sanpham:
                indexa = file_data.loc[file_data['Tên sản phẩm'] == check_sp].index.to_list()[0]
                file_data_id = file_data.loc[indexa]
                dicta = file_data_id.to_dict()             
                for key, value in dicta.items():
                    window[key].update(value)                    
    if event == "Delete":
        if values['Tên sản phẩm'] == '':
            pass
        else :
            check_sp = values['Tên sản phẩm']
            if check_sp in sanpham: 
                indexa = file_data.loc[file_data['Tên sản phẩm'] == check_sp].index.to_list()[0]
                delete_file_data = file_data.drop(indexa)
                delete_file_data.to_excel('D:\\Me\\Python\\Excel\\Data.xlsx', index=False)
                ps.Popup("Xóa thành công")
                clear_input()
    if event == "Data" :
        table_data = file_data.values.tolist()
        window['Table'].update(values = table_data)
 