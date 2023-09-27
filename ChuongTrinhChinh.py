import tkinter as tk
import tkinter.ttk as ttk
import openpyxl
import difflib
import pandas as pd
from datetime import datetime
from PIL import Image, ImageTk

class VnEnDictionary:
    def __init__(self):
        #Khởi tạo đối tượng từ điển và danh sách rỗng để lưu kết quả tìm kiếm
        self.dictionary = self.load_dict_from_excel()
        self.log_results = []

    def __del__(self):
        #Ghi/ lưu từ điển vào tệp Excel khi thoát chương trình
        df = pd.DataFrame(self.dictionary)
        df.to_excel('TuDien_VietAnh23460.xlsx', index = False)

    def load_dict_from_excel(self):
        #Đọc từ điển từ tệp Excel và chuyển thành danh sách
        dictionary = []
        try:
            df = pd.read_excel('TuDien_VietAnh23460.xlsx')
            dictionary = df.to_dict(orient = 'records')
        except Exception as e:
            print("Lỗi khi đọc file Excel:", str(e))
        return dictionary

    def create_menu(self):
        #Tạo cửa số menu chính, chứa các nút để truy cập các chức năng khác nhau
        self.menu_window = tk.Tk()
        self.menu_window.title("Từ Điển Việt-Anh của nhóm 07")

        #Thêm ảnh nền vào cửa sổ menu
        background_image = Image.open("Astronaut.jpg")
        background_photo = ImageTk.PhotoImage(background_image)
        background_label = tk.Label(self.menu_window, image=background_photo)
        background_label.place(relwidth=1, relheight=1)

        #Tạo một khung để chứa các thành phần giao diện và thiết lập kích thước
        self.center_window(self.menu_window, 600, 500)

        #Tạo một nhãn và các ô chọn để người dùng chọn
        menu_label = tk.Label(self.menu_window, text="Chọn một tùy chọn:")
        menu_label.pack(pady=50)

        search_button = tk.Button(self.menu_window, text="Tra từ", command=self.open_search)
        search_button.pack(pady=10)

        add_button = tk.Button(self.menu_window, text="Thêm từ mới", command=self.open_add)
        add_button.pack(pady=10)

        exit_button = tk.Button(self.menu_window, text="Thoát", fg="red", command=self.menu_window.destroy)
        exit_button.pack(pady=10)

        self.menu_window.mainloop()

    def center_window(self, window, width, height):
        #Căn giữa cửa sổ theo kích thước cho trước (chiều rộng x chiều cao)
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    def open_search(self):
        #Phương thức để chuyển từ menu chính đến giao diện có chức năng tra từ
        self.menu_window.destroy()
        self.search_display()

    def open_add(self):
        #Phương thức để chuyển từ menu chính đến giao diện có chức năng thêm từ mới
        self.menu_window.destroy()
        self.add_display()

    def search_display(self):
        #Tạo giao diện cho chức năng tra từ

        #Tạo một cửa sổ mới 'self.root' và thiết lập tiêu đề
        self.root = tk.Tk()
        self.root.title("Tra từ")

        #Thêm ảnh nền vào cửa sổ tra từ
        background_image = Image.open("Galaxy.jpg")
        background_photo = ImageTk.PhotoImage(background_image)
        background_label = tk.Label(self.root, image=background_photo)
        background_label.place(relwidth=1, relheight=1)

        #Tạo một khung để chứa các thành phần giao diện và thiết lập kích thước
        self.center_window(self.root, 800, 600)

        #Tạo một nhãn và ô nhập để người dùng nhập từ cần tra
        search_frame = tk.Frame(self.root)
        search_frame.pack(padx=50, pady=50)

        search_label = tk.Label(search_frame, text="Tra từ:")
        search_label.pack(anchor='w', pady=5)
        self.search_entry = tk.Entry(search_frame)
        self.search_entry.pack(fill='x', padx=5, pady=5)

        #Tạo nút "Tra" để tìm kiếm và nút "Quay lại Menu" để quay lại menu chính
        search_button = tk.Button(search_frame, text="Tra", command=self.search_word)
        search_button.pack(anchor='center', pady=5)

        back_button = tk.Button(search_frame, text="Quay lại Menu", command=self.back_to_menu)
        back_button.pack(anchor='center', pady=5)

        #Tạo một nhãn để hiển thị kết quả tìm kiếm
        self.result_label = tk.Label(self.root, text="", justify='left')
        self.result_label.pack()

        self.root.mainloop()

    def add_display(self):
        #Tạo giao diện cho chức năng thêm từ mới

        #Tạo một cửa sổ mới 'self.root' và thiết lập tiêu đề
        self.root = tk.Tk()
        self.root.title("Thêm từ mới")

        #Thêm ảnh nền vào cửa sổ thêm từ mới
        background_image = Image.open("Galaxy.jpg")
        background_photo = ImageTk.PhotoImage(background_image)
        background_label = tk.Label(self.root, image=background_photo)
        background_label.place(relwidth=1, relheight=1)

        #Tạo một khung để chứa các thành phần giao diện và thiết lập kích thước
        self.center_window(self.root, 800, 600)

        #Tạo nhãn và ô nhập để người dùng nhập từ mới và nghĩa của nó
        add_frame = tk.Frame(self.root)
        add_frame.pack(padx=40, pady=40)

        add_label = tk.Label(add_frame, text="Thêm từ mới:")
        add_label.pack(anchor='w', pady=5)
        self.new_word_entry = tk.Entry(add_frame)
        self.new_word_entry.pack(fill='x', padx=5, pady=5)

        meaning_label = tk.Label(add_frame, text="Nghĩa tiếng Anh:")
        meaning_label.pack(anchor='w', pady=5)
        self.meaning_entry = tk.Entry(add_frame)
        self.meaning_entry.pack(fill='x', padx=5, pady=5)

        #Tạo nút "Thêm" để thêm từ mới và nút "Quay lại Menu" để quay lại menu chính
        add_button = tk.Button(add_frame, text="Thêm", command=self.add_new_word)
        add_button.pack(anchor='center', pady=5)

        back_button = tk.Button(add_frame, text="Quay lại Menu", command=self.back_to_menu)
        back_button.pack(anchor='center', pady=5)

        self.result_label = tk.Label(self.root, text="", justify='left')
        self.result_label.pack()

        self.root.mainloop()

    def back_to_menu(self):
        #Quay lại menu chính từ bất kỳ chức năng nào
        self.root.destroy()
        self.create_menu()

    def search_word(self):
        #Tìm kiếm từ trong từ điển

        #Lấy thời gian hiện tại và từ cần tra từ ô nhập
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        word = self.search_entry.get().strip()
        print(word)
        success = False

        #Duyệt qua từ điển và nếu từ được tìm thấy, hiển thị nghĩa của từ và ghi kết quả ra file log
        for w in self.dictionary:
            if w['word'] == word.lower():
                meaning = w['meaning']
                self.log_results.append(f"Tra từ: {word}\n{meaning}")
                self.result_label.config(text = meaning)
                success = True
                with open("log.txt", "a", encoding = "utf-8") as log_file:
                    log_file.write(f"[{timestamp}]: Tra cứu từ: {word}\n=>> Nghĩa: {meaning}\n")
                break

        #Nếu không tìm thấy, ghi kết quả ra file log và đề xuất các từ gần giống
        if not success:
            with open("log.txt", "a", encoding="utf-8") as log_file:
                log_file.write(f"[{timestamp}]: Không tìm thấy từ: {word}\n")
            suggestions = self.get_word_suggestions(word)
            if suggestions:
                self.result_label.config(text = f"Từ {word} không tồn tại.\nCó thể bạn muốn tìm: {', '.join(suggestions)}")
            else:
                self.result_label.config(text = f"Từ {word} không tồn tại!")

    def get_word_suggestions(self, word):
        #Đề xuất các từ gần giống với từ cần tra (khi hiện kết quả không có từ trong từ điển)
        words = [
            w['word'] for w in self.dictionary if isinstance(w['word'], str)]
        suggestions = difflib.get_close_matches(
            word = word.lower(), possibilities=words, n = 5, cutoff = 0.7)
        return suggestions

    def add_new_word(self):
        #Thêm từ mới vào từ điển
        
        #Lấy thời gian hiện tại, từ mới và nghĩa của từ
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        #Tạo kết quả trả về trong app nếu người dùng không nhập 1 trong 2 mục
        new_word = self.new_word_entry.get().strip().lower()
        if not new_word:
            self.result_label.config(text = "Vui lòng nhập từ cần thêm!")
            return

        meaning = self.meaning_entry.get().strip()
        if not meaning:
            self.result_label.config(text = "Vui lòng nhập nghĩa của từ!")
            return

        #Kiểm tra xem từ đã tồn tại trong từ điển chưa, nếu từ chưa tồn tại thì thêm nó vào từ điển và ghi ra file log
        if new_word in [w['word'] for w in self.dictionary if isinstance(w['word'], str)]:
            self.result_label.config(
                text = "Từ này đã tồn tại, vui lòng nhập từ khác!")
        else:
            self.dictionary.append({'word': new_word, 'meaning': meaning})
            self.log_results.append(f"Thêm từ mới: {new_word} - Nghĩa: {meaning}")

            with open("log.txt", "a", encoding = "utf-8") as log_file:
                log_file.write(f"[{timestamp}]: Thêm từ mới: {new_word}\n=>> Nghĩa: {meaning}\n")

            self.result_label.config(text = f"Đã thêm từ mới: {new_word}")

    def menu(self):
        #Bắt đầu chương trình bằng cách tạo cửa sổ menu chính
        self.create_menu()

"""Kiểm tra xem chương trình có được chạy trực tiếp hay là được import vào một chương trình khác
và gọi hàm menu() để bắt đầu chương trình"""
if __name__ == "__main__":
    VnEnDictionary().menu()