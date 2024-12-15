import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
from datetime import datetime

root = tk.Tk()
root.title("Thông tin nhân viên")
root.geometry("800x400")

def save_to_csv():
    data = [
        entry_ma.get(),
        entry_ten.get(),
        combo_don_vi.get(),
        combo_chuc_danh.get(),
        "Nam" if radio_gender.get() == "Nam" else "Nữ",
        entry_ngay_sinh.get(),
        entry_cmnd.get(),
        entry_ngay_cap.get(),
        entry_noi_cap.get(),
    ]
    if "" in data:
        messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
        return

    with open("nhanvien.csv", mode="a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(data)
    messagebox.showinfo("Thành công", "Dữ liệu đã được lưu!")
    clear_entries()

def clear_entries():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    combo_don_vi.set("")
    combo_chuc_danh.set("")
    entry_ngay_sinh.delete(0, tk.END)
    entry_cmnd.delete(0, tk.END)
    entry_ngay_cap.delete(0, tk.END)
    entry_noi_cap.delete(0, tk.END)
    radio_gender.set("")

def show_today_birthdays():
    today = datetime.now().strftime("%d/%m")
    try:
        with open("nhanvien.csv", mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            birthdays = [
                row for row in reader
                if len(row) > 5 and today in row[5]
            ]
        if birthdays:
            result = "\n".join([" | ".join(row) for row in birthdays])
            messagebox.showinfo("Sinh nhật hôm nay", result)
        else:
            messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "File dữ liệu chưa tồn tại!")

def export_to_excel():
    try:
        with open("nhanvien.csv", mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            sorted_data = sorted(reader, key=lambda x: datetime.strptime(x[5], "%d/%m/%Y"), reverse=True)

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerows(sorted_data)
            messagebox.showinfo("Thành công", "Danh sách đã được xuất!")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "File dữ liệu chưa tồn tại!")
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))

frame = ttk.Frame(root, padding=10)
frame.pack(fill=tk.BOTH, expand=True)

ttk.Label(frame, text="Mã *").grid(row=0, column=0, sticky=tk.W, pady=5)
entry_ma = ttk.Entry(frame)
entry_ma.grid(row=0, column=1, pady=5)

ttk.Label(frame, text="Tên *").grid(row=0, column=2, sticky=tk.W, pady=5)
entry_ten = ttk.Entry(frame)
entry_ten.grid(row=0, column=3, pady=5)

ttk.Label(frame, text="Đơn vị *").grid(row=1, column=0, sticky=tk.W, pady=5)
combo_don_vi = ttk.Combobox(frame, values=["Phân xưởng", "Văn phòng", "Ban quản lý"])
combo_don_vi.grid(row=1, column=1, pady=5)

ttk.Label(frame, text="Chức danh *").grid(row=1, column=2, sticky=tk.W, pady=5)
combo_chuc_danh = ttk.Combobox(frame, values=["Nhân viên", "Quản lý", "Giám đốc"])
combo_chuc_danh.grid(row=1, column=3, pady=5)

ttk.Label(frame, text="Giới tính *").grid(row=2, column=0, sticky=tk.W, pady=5)
radio_gender = tk.StringVar()
ttk.Radiobutton(frame, text="Nam", variable=radio_gender, value="Nam").grid(row=2, column=1, sticky=tk.W)
ttk.Radiobutton(frame, text="Nữ", variable=radio_gender, value="Nữ").grid(row=2, column=2, sticky=tk.W)

ttk.Label(frame, text="Ngày sinh (DD/MM/YYYY) *").grid(row=3, column=0, sticky=tk.W, pady=5)
entry_ngay_sinh = ttk.Entry(frame)
entry_ngay_sinh.grid(row=3, column=1, pady=5)

ttk.Label(frame, text="Số CMND *").grid(row=3, column=2, sticky=tk.W, pady=5)
entry_cmnd = ttk.Entry(frame)
entry_cmnd.grid(row=3, column=3, pady=5)

ttk.Label(frame, text="Ngày cấp (DD/MM/YYYY)").grid(row=4, column=0, sticky=tk.W, pady=5)
entry_ngay_cap = ttk.Entry(frame)
entry_ngay_cap.grid(row=4, column=1, pady=5)

ttk.Label(frame, text="Nơi cấp").grid(row=4, column=2, sticky=tk.W, pady=5)
entry_noi_cap = ttk.Entry(frame)
entry_noi_cap.grid(row=4, column=3, pady=5)

btn_save = ttk.Button(frame, text="Lưu thông tin", command=save_to_csv)
btn_save.grid(row=5, column=0, pady=10)

btn_today_birthdays = ttk.Button(frame, text="Sinh nhật ngày hôm nay", command=show_today_birthdays)
btn_today_birthdays.grid(row=5, column=1, pady=10)

btn_export = ttk.Button(frame, text="Xuất toàn bộ danh sách", command=export_to_excel)
btn_export.grid(row=5, column=2, pady=10)

btn_clear = ttk.Button(frame, text="Xóa trắng", command=clear_entries)
btn_clear.grid(row=5, column=3, pady=10)

root.mainloop()