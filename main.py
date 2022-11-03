
import pandas as pd
import PySimpleGUI as ps

# Xuất data từ file Excel
file_data = pd.read_excel('D:\\Me\\Python\\Du_an\\Data.xlsx')
table_data = file_data.values.tolist()
table_header = file_data.columns.tolist()
id_number = file_data['ID'].values
sanpham = file_data['Tên sản phẩm'].values

# Tạo Bảng điểu khiển
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
    [ps.Table(values = table_data,    
            headings = table_header,
            key="Table",
            row_height = 30,
            justification = 'center',
            expand_x= True,
            expand_y= True,)]
    ]


window = ps.Window("Thông tin sản phẩm",layout)
# Tạo hàm clear
def clear_input():
    for key in values:
        window[key](" ")
    return None
    

while True:
    # Lấy giá trị button và giá trị truyền vào
    event,values = window.read()
    # Đóng window
    if event == ps.WIN_CLOSED or event == "Exit":
        break
    # Thêm, sửa sản phẩm 
    if event == "Save":
        check_id = int(values['ID'])
        check_sp = values['Tên sản phẩm']
        if check_id == "" or check_sp == "":
            ps.popup("Vui lòng nhập đầy đủ thông tin sản phẩm")
        elif check_id in id_number:
            ps.popup('Một sản phẩm không thể có 2 ID')
        # Thêm sản phẩm mới
        # sửa sản phẩm theo tên sản phẩm
        elif check_sp in sanpham:
            indexa = file_data.loc[file_data['Tên sản phẩm'] == check_sp].index.to_list()[0]
            header_list = list(file_data.columns.values)
            for key in header_list : 
                file_data.loc[indexa,key] = values[key]
            file_data.to_excel('D:\\Me\\Python\\Du_an\\Data.xlsx', index=False)
            ps.Popup("Sửa thành công")
            clear_input()
        else :
            del values['Table']
            file_data = file_data.append(values,ignore_index=True)
            file_data.to_excel('D:\\Me\\Python\\Du_an\\Data.xlsx', index=False)
            ps.Popup("Thêm thành công")
            clear_input()
    # Tìm sản phẩm theo ID vầ Tên sản phẩm 
    if event == "Search":
        check_sp = values['Tên sản phẩm']
        if values['ID'] == '':
            check_id = 0
        else: check_id = int(values['ID'])
        if check_id == 0 and check_sp =='' :
            ps.popup('Vui lòng nhập thông tin sản phẩm ')
        else :
            # Tìm theo tên sản phẩm và xuất giá trị vào các trường
            if check_sp in sanpham and check_id == 0:
                indexa = file_data.loc[file_data['Tên sản phẩm'] == check_sp].index.to_list()[0]
                file_data_id = file_data.loc[indexa]
                dicta = file_data_id.to_dict()             
                for key, value in dicta.items():
                    window[key].update(value)
            # Tìm theo tên ID và xuất giá trị vào các trường      
            elif check_id in id_number:
                indexa = file_data.loc[file_data['ID'] == int(check_id)].index.to_list()[0]
                file_data_id = file_data.loc[indexa]
                dicta = file_data_id.to_dict()             
                for key, value in dicta.items():
                    window[key].update(value)
            else: ps.popup('Sản phẩm chưa tồn tại')  
    # Xóa sản phẩm              
    if event == "Delete":
        if values['Tên sản phẩm'] == '':
            pass
        else :
            check_sp = values['Tên sản phẩm']
            if check_sp in sanpham: 
                indexa = file_data.loc[file_data['Tên sản phẩm'] == check_sp].index.to_list()[0]
                delete_file_data = file_data.drop(indexa)
                delete_file_data.to_excel('D:\\Me\\Python\\Du_an\\Data.xlsx', index=False)
                ps.Popup("Xóa thành công")
                clear_input()
    # Clear bảng tính
    if event == "Reset" :
        clear_input()
    # Xuất data sau khi sửa 
    if event == "Data" :
        val = file_data.values.tolist()
        window['Table'].update(values = val)    
 