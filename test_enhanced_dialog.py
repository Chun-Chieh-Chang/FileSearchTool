import tkinter as tk
from tkinter import ttk, filedialog
import os

def test_enhanced_folder_dialog():
    root = tk.Tk()
    root.geometry("400x250")
    root.title("測試增強版資料夾選擇")
    
    folder_path = tk.StringVar()
    
    def browse_folder():
        try:
            print("開啟資料夾選擇...")
            initial_dir = folder_path.get() if folder_path.get() and os.path.exists(folder_path.get()) else None
            
            foldername = filedialog.askdirectory(
                title="選擇要搜尋的資料夾",
                initialdir=initial_dir,
                mustexist=True
            )
            
            if foldername:
                if os.path.exists(foldername) and os.path.isdir(foldername):
                    try:
                        os.listdir(foldername)
                        folder_path.set(foldername)
                        status_label.config(text=f"成功選擇: {foldername}")
                        print(f"成功選擇資料夾: {foldername}")
                    except PermissionError:
                        status_label.config(text="權限不足")
                        print("權限不足")
                    except Exception as e:
                        status_label.config(text=f"訪問錯誤: {e}")
                        print(f"訪問錯誤: {e}")
                else:
                    status_label.config(text="無效路徑")
                    print("無效路徑")
            else:
                status_label.config(text="已取消選擇")
                print("已取消選擇")
                
        except Exception as e:
            error_msg = f"對話框錯誤: {e}"
            status_label.config(text=error_msg)
            print(error_msg)
    
    frame = ttk.Frame(root, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)
    
    ttk.Label(frame, text="增強版資料夾選擇測試").pack(pady=10)
    
    entry = ttk.Entry(frame, textvariable=folder_path, width=50)
    entry.pack(pady=5, fill=tk.X)
    
    ttk.Button(frame, text="瀏覽資料夾", command=browse_folder).pack(pady=10)
    
    status_label = ttk.Label(frame, text="點擊按鈕測試", foreground="blue")
    status_label.pack(pady=5)
    
    print("測試視窗已開啟，請點擊瀏覽按鈕...")
    root.mainloop()

if __name__ == "__main__":
    test_enhanced_folder_dialog()