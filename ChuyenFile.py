#Import thư viện pandas với tên viết tắt là pd
import pandas as pd

#Đọc dữ liệu từ file 'TuDien_VietAnh23460.txt'
with open('TuDien_VietAnh23460.txt', 'r', encoding='utf-8') as file:
    f = file.read()  #Đọc toàn bộ nội dung của tệp và lưu vào biến f
    words = f.split('\n@')  #Tách nội dung thành danh sách các từ với dấu phân cách là '\n@'

data = []  #Khởi tạo một danh sách rỗng để lưu dữ liệu

#Lặp qua từng từ trong danh sách từ
for word in words:
    w = {}  #Khởi tạo một từ điển rỗng để lưu thông tin về từ và nghĩa
    w['word'] = word.split('\n', 1)[0]  #Tách từ đầu tiên trong từ và lưu vào key 'word'
    
    try:
        w['meaning'] = word.split('\n', 1)[1]  #Tách nghĩa và lưu vào key 'meaning'
    except:
        w['meaning'] = 'No data'  #Nếu không có nghĩa, gán giá trị 'No data' cho key 'meaning'
    
    data.append(w)  #Thêm từ điển w vào danh sách data

#Bỏ qua những từ không phải là string
data = [w for w in data if isinstance(w['word'], str)]

#Chuyển danh sách dữ liệu thành DataFrame bằng thư viện pandas
df = pd.DataFrame(data)

#Lưu DataFrame thành file Excel tên 'TuDien_VietAnh23460.xlsx'
df.to_excel('TuDien_VietAnh23460.xlsx', index=False)