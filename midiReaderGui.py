import tkinter as tk
from tkinter import colorchooser,filedialog,font,messagebox
import openpyxl as op
import midiReader as mr
import traceback
from functools import wraps

class midiReaderGui:
    try:
        def __init__(self,master):
            self.master=master
            self.master.title("MIDI Reader")
            self.master.geometry("500x150")

            #excel生成用に渡す引数
            self.open_dir=None
            self.save_dir=None
            self.col_dir=None
            check=tk.BooleanVar()
            #表示する色の初期化
            
            def show_on_error(function):
                "関数実行中に例外が発生したら、showerrorするデコレータ"
                @wraps(function)
                def show_error(*args,**kwargs):
                    try:
                        function(*args,**kwargs)
                    except Exception as e:
                        print(traceback.format_exc())
                        title = e.__class__.__name__
                        message = traceback.format_exc()
                        messagebox.showerror(f"{title}",
                                            f"{message}")
                return show_error

            ###event処理達
            @show_on_error
            def generate_event():
                #excelシートの生成
                wb=op.Workbook()
                wb.save(self.save_dir)
                if check:
                    m=mr.midiReader(self.open_dir,self.save_dir,self.bpm_entry.get(),float(self.delay_entry.get()))
                else:
                    m=mr.midiReader(self.open_dir,self.save_dir,self.bpm_entry.get(),float(self.delay_entry.get()))

            @show_on_error
            def open_select_file():
                file_type = [("読込midi File", "*.mid")]
                selected_file = filedialog.askopenfilename(filetypes=file_type,defaultextension="xlsx")
                if selected_file:
                    # for file in selected_files:
                    #     self.listbox_files.insert(tk.END, file)
                    self.open_dir=selected_file
                    self.open_dir_label.config(text=selected_file)
            
            @show_on_error
            def open_color_file():
                file_type = [("読込Excel", "*.xlsx")]
                selected_file = filedialog.askopenfilename(filetypes=file_type,defaultextension="xlsx")
                if selected_file:
                    # for file in selected_files:
                    #     self.listbox_files.insert(tk.END, file)
                    self.col_dir=selected_file
                    self.open_color_label.config(text=selected_file)
            
            @show_on_error
            def save_select_file():
                file_type = [("保存Excel", "*.xlsx")]
                selected_file = filedialog.asksaveasfilename(filetypes=file_type,defaultextension="xlsx")
                if selected_file:
                    # for file in selected_files:
                    #     self.listbox_files.insert(tk.END, file)
                    self.save_dir=selected_file
                    self.save_dir_label.config(text=selected_file)
                    

            # #gridの行列はこれで取得できる(デバッグ用)
            # def callback(event):
            #     info = event.widget.grid_info()
            #     print(info['row'],info['column'])

            #widget
            
            #label
            self.bpm_label=tk.Label(text="BPM")
            self.open_label=tk.Label(text="MIDIファイル")
            self.save_label=tk.Label(text="保存先")
            self.open_dir_label=tk.Label(text="読み込むファイルを選択(.mid)")
            self.save_dir_label=tk.Label(text="保存ファイルを選択(.xlsx)")
            self.delay_label=tk.Label(text="Delay(sec)")
            #entry
            self.bpm_entry=tk.Entry(justify="center")
            self.delay_entry=tk.Entry(justify="center")

            #button
            self.open_selectFile_button=tk.Button(text="参照",command=open_select_file)
            self.save_selectFile_button=tk.Button(text="参照",command=save_select_file)
            self.generate_button=tk.Button(text="generate",command=generate_event)
            #scroll,listbox

            #gridで配置
            self.open_label.grid(row=0,column=0,sticky="e")
            self.save_label.grid(row=1,column=0,sticky="e")
            self.bpm_label.grid(row=2,column=0,sticky="e")
            self.delay_label.grid(row=3,column=0,sticky="e")
            self.open_selectFile_button.grid(row=0,column=1,sticky="e")
            self.save_selectFile_button.grid(row=1,column=1,sticky="e")
            self.open_dir_label.grid(row=0,column=2,sticky="w")
            self.save_dir_label.grid(row=1,column=2,sticky="w")
            self.bpm_entry.grid(row=2,column=1,sticky="w")
            self.delay_entry.grid(row=3,column=1,sticky="w")
            self.generate_button.grid(row=4,column=0,sticky="w",padx=10,pady=10)

            self.master.grid_columnconfigure(0,weight=1,minsize=100)    
            self.master.grid_columnconfigure(1,weight=1,minsize=100)    
            self.master.grid_columnconfigure(2,weight=1,minsize=100)    
            self.master.grid_columnconfigure(3,weight=1,minsize=100)    
            self.master.grid_rowconfigure(7,weight=1,minsize=100)

            # self.master.bind_all("<1>",callback) #gridの行列取得
            
            self.delay_entry.insert(0,"0")
    except Exception as e:
        print("aaaaaaaaaaaaa")
        print(traceback.format_exc())
        title = e.__class__.__name__
        message = traceback.format_exc()
        messagebox.showerror(f"{title}",
                                f"{message}")
        


if __name__ == '__main__':
    root = tk.Tk()
    make = midiReaderGui(root)
    root.mainloop()       
        