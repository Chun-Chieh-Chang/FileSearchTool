import tkinter as tk
from tkinter import ttk, filedialog

def test_folder_dialog():
    root = tk.Tk()
    root.geometry("300x200")
    root.title("測試資料夾選擇")
    
    folder_path = tk.StringVar()
    
    def browse_folder():
        print("開啟資料夾選擇...")
        foldername = filedialog.askdirectory()
        print(f"選擇結果: {foldername}")
        if foldername:
            folder_path.set(foldername)
            status_label.config(text=f"已選擇: {foldername}")
        else:
            status_label.config(text="未選擇資料夾")
    
    frame = ttk.Frame(root, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)
    
    ttk.Label(frame, text="測試資料夾選擇功能").pack(pady=10)
    
    entry = ttk.Entry(frame, textvariable=folder_path, width=40)
    entry.pack(pady=5, fill=tk.X)
    
    ttk.Button(frame, text="瀏覽資料夾", command=browse_folder).pack(pady=10)
    
    status_label = ttk.Label(frame, text="點擊按鈕測試")
    status_label.pack(pady=5)
    
    print("測試視窗已開啟，請點擊瀏覽按鈕...")
    root.mainloop()

if __name__ == "__main__":
    test_folder_dialog()