import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
# --- Общие настройки (используются ТОЛЬКО в режимах склейки результатов) ---
SHEETS = [
   "visit order",
   "dropped locations",
   "visit duration",
   "visits per week",
   "employees",
   "data",
   "routes_metrics",
]
VARIABLE_SHEETS = ["visit order", "visit duration", "visits per week"]
FIXED_SHEETS = ["dropped locations", "employees", "data", "routes_metrics"]
EMPLOYEE_SHEETS = [
   "visit order",
   "visit duration",
   "visits per week",
   "data",
   "routes_metrics",
]
ID_SHEETS = ["employees", "dropped locations"]

class MainApp:
   def __init__(self, root):
       self.root = root
       self.root.title("Выбор режима склейки Excel-файлов")
       self.root.geometry("550x300")
       self.root.resizable(False, False)
       # === Группа: Склейка результатов ===
       tk.Label(
           root,
           text="Склейка результатов",
           font=("Arial", 12, "bold"),
           fg="#2196F3",
       ).pack(pady=(10, 0))
       tk.Button(
           root,
           text="Обычная склейка (несколько файлов)",
           font=("Arial", 10),
           command=self.open_merge_app,
           bg="#2196F3",
           fg="white",
           padx=20,
           pady=5,
       ).pack(pady=5, padx=60, fill="x")
       tk.Button(
           root,
           text="Взаимоисключающая склейка (два файла)",
           font=("Arial", 10),
           command=self.open_exclude_app,
           bg="#F44336",
           fg="white",
           padx=20,
           pady=5,
       ).pack(pady=5, padx=60, fill="x")
       # === Группа: Склейка импорта ===
       tk.Label(
           root, text="Склейка импорта", font=("Arial", 12, "bold"), fg="#4CAF50"
       ).pack(pady=(15, 0))
       tk.Button(
           root,
           text="Простая склейка импорта (одинаковые файлы)",
           font=("Arial", 10),
           command=self.open_import_app,
           bg="#4CAF50",
           fg="white",
           padx=20,
           pady=5,
       ).pack(pady=5, padx=60, fill="x")
   def open_merge_app(self):
       """Открыть окно обычной склейки"""
       self.root.withdraw()
       merge_window = tk.Toplevel(self.root)
       merge_window.protocol("WM_DELETE_WINDOW", lambda: self.on_closing(merge_window))
       ExcelMergerApp(merge_window, self.root)
   def open_exclude_app(self):
       """Открыть окно взаимоисключающей склейки"""
       self.root.withdraw()
       exclude_window = tk.Toplevel(self.root)
       exclude_window.protocol(
           "WM_DELETE_WINDOW", lambda: self.on_closing(exclude_window)
       )
       ExcelExclusionApp(exclude_window, self.root)
   def open_import_app(self):
       """Открыть окно простой склейки импорта"""
       self.root.withdraw()
       import_window = tk.Toplevel(self.root)
       import_window.protocol(
           "WM_DELETE_WINDOW", lambda: self.on_closing(import_window)
       )
       ExcelImportMergerApp(import_window, self.root)
   def on_closing(self, window):
       """Закрытие дочернего окна и возврат в главное"""
       window.destroy()
       self.root.deiconify()

# --- Режим 1: Обычная склейка (для результатов) ---
class ExcelMergerApp:
   def __init__(self, root, parent_root):
       self.root = root
       self.parent_root = parent_root
       self.root.title("Обычная склейка Excel-файлов")
       self.root.geometry("500x200")
       self.file_paths = []
       self.label = tk.Label(
           root, text="Выберите файлы для склейки", font=("Arial", 12)
       )
       self.label.pack(pady=20)
       self.select_button = tk.Button(
           root, text="Выбрать файлы", command=self.select_files, font=("Arial", 10)
       )
       self.select_button.pack(pady=5)
       self.progress = ttk.Progressbar(
           root, orient="horizontal", length=400, mode="determinate"
       )
       self.progress.pack(pady=10)
       self.status_label = tk.Label(root, text="", font=("Arial", 10))
       self.status_label.pack(pady=5)
   def select_files(self):
       self.file_paths = filedialog.askopenfilenames(
           title="Выберите Excel-файлы", filetypes=[("Excel files", "*.xlsx")]
       )
       if self.file_paths:
           self.select_button.config(state="disabled")
           self.label.config(text=f"Выбрано файлов: {len(self.file_paths)}")
           threading.Thread(target=self.merge_files, daemon=True).start()
       else:
           self.label.config(text="Файлы не выбраны")
   def merge_files(self):
       try:
           if not self.file_paths:
               raise Exception("Не выбрано ни одного файла")
           merged_data = {sheet: [] for sheet in SHEETS}
           unified_columns = {}
           missing_column_warnings = []
           total_steps = len(self.file_paths) * len(SHEETS)
           current_step = 0
           # --- Шаг 1: Определение единой шапки для VARIABLE_SHEETS ---
           for sheet in VARIABLE_SHEETS:
               all_columns = set()
               file_columns = {}
               for file_path in self.file_paths:
                   try:
                       with pd.ExcelFile(file_path) as xls:
                           sheet_names_lower = {
                               name.strip().lower(): name for name in xls.sheet_names
                           }
                           matched_name = sheet_names_lower.get(sheet)
                           if not matched_name:
                               continue
                           df = pd.read_excel(xls, sheet_name=matched_name, header=1)
                           df.dropna(how="all", inplace=True)
                           cols = df.columns.tolist()
                           file_columns[file_path] = cols
                           all_columns.update(cols)
                   except Exception as e:
                       print(f"[INFO] Ошибка при чтении '{sheet}' из {file_path}: {e}")
                       file_columns[file_path] = []
               # Базовый файл — с наибольшим количеством колонок
               base_file = max(
                   file_columns.items(),
                   key=lambda x: len(x[1]) if x[1] else 0,
                   default=(None, []),
               )[0]
               if base_file and file_columns[base_file]:
                   unified_columns[sheet] = file_columns[base_file][:]
                   for col in all_columns:
                       if col not in unified_columns[sheet]:
                           unified_columns[sheet].append(col)
               else:
                   unified_columns[sheet] = list(all_columns)
               # Проверка недостающих колонок
               ref_cols = set(unified_columns[sheet])
               for fp, cols in file_columns.items():
                   if not cols:
                       missing_column_warnings.append(
                           f"Файл '{os.path.basename(fp)}': нет данных на листе '{sheet}'"
                       )
                       continue
                   missing = ref_cols - set(cols)
                   for col in missing:
                       missing_column_warnings.append(
                           f"Файл '{os.path.basename(fp)}', лист '{sheet}': отсутствует колонка '{col}'"
                       )
           current_step += len(VARIABLE_SHEETS) * len(self.file_paths)
           self.update_progress(int(current_step / total_steps * 100))
           # --- Шаг 2: Чтение и объединение ---
           for file_path in self.file_paths:
               try:
                   with pd.ExcelFile(file_path) as xls:
                       sheet_names_lower = {
                           name.strip().lower(): name for name in xls.sheet_names
                       }
                       for sheet in SHEETS:
                           matched_name = sheet_names_lower.get(sheet)
                           if not matched_name:
                               continue
                           try:
                               df = pd.read_excel(
                                   xls, sheet_name=matched_name, header=1
                               )
                               df.dropna(how="all", inplace=True)
                               if df.empty:
                                   continue
                               df = df.reset_index(drop=True)
                               if sheet in VARIABLE_SHEETS:
                                   target_cols = unified_columns[sheet]
                                   aligned_df = pd.DataFrame(
                                       {
                                           col: df[col] if col in df.columns else None
                                           for col in target_cols
                                       }
                                   )
                                   df = aligned_df
                               merged_data[sheet].append(df)
                           except Exception as e:
                               print(f"[ERROR] Чтение '{sheet}' из {file_path}: {e}")
               except Exception as e:
                   print(f"[ERROR] Открытие файла {file_path}: {e}")
               current_step += len(FIXED_SHEETS)
               self.update_progress(int(current_step / total_steps * 100))
           # --- Шаг 3: Финальные данные ---
           final_dfs = {}
           for sheet in SHEETS:
               if merged_data[sheet]:
                   final_dfs[sheet] = pd.concat(merged_data[sheet], ignore_index=True)
               else:
                   header_columns = None
                   for fp in self.file_paths:
                       try:
                           with pd.ExcelFile(fp) as xls:
                               sn = {n.strip().lower(): n for n in xls.sheet_names}
                               mn = sn.get(sheet)
                               if mn:
                                   df_h = pd.read_excel(
                                       xls, sheet_name=mn, nrows=0, header=1
                                   )
                                   header_columns = df_h.columns.tolist()
                                   break
                       except Exception:
                           continue
                   final_dfs[sheet] = (
                       pd.DataFrame(columns=header_columns)
                       if header_columns
                       else pd.DataFrame()
                   )
           # --- Шаг 4: Сохранение ---
           output_path = os.path.join(
               os.path.dirname(self.file_paths[0]), "merged_result.xlsx"
           )
           with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
               for sheet, df in final_dfs.items():
                   if not df.empty or len(df.columns) > 0:
                       df.to_excel(
                           writer,
                           sheet_name=sheet,
                           startrow=1,
                           index=False,
                           header=True,
                       )
                   else:
                       writer.book.create_sheet(sheet)
           self.update_progress(100)
           if missing_column_warnings:
               warn_msg = "Обнаружены различия в структуре таблиц:\n" + "\n".join(
                   missing_column_warnings[:50]
               )
               if len(missing_column_warnings) > 50:
                   warn_msg += f"\n...и ещё {len(missing_column_warnings) - 50} предупреждений."
               self.root.after(
                   0, lambda: messagebox.showwarning("Предупреждение", warn_msg)
               )
           self.root.after(0, lambda: self.show_completion(output_path))
       except Exception as e:
           self.root.after(
               0, lambda: messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")
           )
   def update_progress(self, value):
       self.root.after(0, lambda: self.progress.config(value=value))
       self.root.after(
           0, lambda: self.status_label.config(text=f"Обработка: {value}%")
       )
   def show_completion(self, path):
       messagebox.showinfo("Готово", f"Файлы успешно склеены!\nФайл: {path}")
       self.root.destroy()
       self.parent_root.deiconify()

# --- Режим 2: Взаимоисключающая склейка (для результатов) ---
class ExcelExclusionApp:
   def __init__(self, root, parent_root):
       self.root = root
       self.parent_root = parent_root
       self.root.title("Взаимоисключающая склейка")
       self.root.geometry("550x250")
       self.main_file = ""
       self.extra_file = ""
       tk.Label(
           root, text="Выберите файл для взаимоисключающей склейки", font=("Arial", 12)
       ).pack(pady=20)
       tk.Button(
           root,
           text="Выбрать файлы",
           command=self.select_files,
           font=("Arial", 10),
           bg="#4CAF50",
           fg="white",
       ).pack(pady=10)
       self.progress = ttk.Progressbar(
           root, orient="horizontal", length=450, mode="determinate"
       )
       self.progress.pack(pady=10)
       self.status_label = tk.Label(root, text="", font=("Arial", 10))
       self.status_label.pack(pady=5)
   def select_files(self):
       self.status_label.config(text="Выбирается основной файл...")
       self.root.update()
       main_file = filedialog.askopenfilename(
           title="Выберите основной файл", filetypes=[("Excel файлы", "*.xlsx")]
       )
       if not main_file:
           messagebox.showwarning("Отмена", "Основной файл не выбран.")
           return
       self.main_file = main_file
       self.status_label.config(text="Выбирается дополнительный файл...")
       self.root.update()
       extra_file = filedialog.askopenfilename(
           title="Выберите дополнительный файл", filetypes=[("Excel файлы", "*.xlsx")]
       )
       if not extra_file:
           messagebox.showwarning("Отмена", "Дополнительный файл не выбран.")
           return
       self.extra_file = extra_file
       threading.Thread(target=self.merge_files, daemon=True).start()
   def merge_files(self):
       try:
           with pd.ExcelFile(self.main_file) as main_xls, pd.ExcelFile(
               self.extra_file
           ) as extra_xls:
               sheet_names_main = {
                   name.strip().lower(): name for name in main_xls.sheet_names
               }
               sheet_names_extra = {
                   name.strip().lower(): name for name in extra_xls.sheet_names
               }
           for sheet in SHEETS:
               if sheet not in sheet_names_main:
                   raise FileNotFoundError(
                       f"Лист '{sheet}' отсутствует в основном файле."
                   )
               if sheet not in sheet_names_extra:
                   raise FileNotFoundError(
                       f"Лист '{sheet}' отсутствует в дополнительном файле."
                   )
           missing_column_warnings = []
           unified_columns = {}
           total_steps = len(SHEETS) * 2
           current_step = 0
           # Анализ шапок
           for sheet in VARIABLE_SHEETS:
               main_name = sheet_names_main[sheet]
               extra_name = sheet_names_extra[sheet]
               df_main = pd.read_excel(
                   self.main_file, sheet_name=main_name, header=1
               ).dropna(how="all")
               df_extra = pd.read_excel(
                   self.extra_file, sheet_name=extra_name, header=1
               ).dropna(how="all")
               main_cols = df_main.columns.tolist()
               extra_cols = df_extra.columns.tolist()
               unified_columns[sheet] = main_cols
               missing = [col for col in main_cols if col not in extra_cols]
               if missing:
                   missing_column_warnings.append(
                       f"Лист '{sheet}': в дополнительном файле отсутствуют колонки: {', '.join(missing)}"
                   )
               current_step += 1
               self.update_progress(int(current_step / total_steps * 50))
           # Чтение всех листов
           main_dfs = {
               sheet: pd.read_excel(
                   self.main_file, sheet_name=sheet_names_main[sheet], header=1
               ).dropna(how="all")
               for sheet in SHEETS
           }
           extra_dfs = {
               sheet: pd.read_excel(
                   self.extra_file, sheet_name=sheet_names_extra[sheet], header=1
               ).dropna(how="all")
               for sheet in SHEETS
           }
           # Новое правило: удаление из dropped locations
           if (
               "id" in main_dfs["dropped locations"].columns
               and "location_id" in extra_dfs["visit order"].columns
           ):
               loc_ids = set(extra_dfs["visit order"]["location_id"].dropna())
               main_dfs["dropped locations"] = main_dfs["dropped locations"][
                   ~main_dfs["dropped locations"]["id"].isin(loc_ids)
               ]
           # Обработка всех листов
           result_data = {}
           for sheet in SHEETS:
               df_main = main_dfs[sheet]
               df_extra = extra_dfs[sheet]
               key_col = "employee" if sheet in EMPLOYEE_SHEETS else "id"
               if key_col not in df_main.columns or key_col not in df_extra.columns:
                   raise KeyError(f"Колонка '{key_col}' отсутствует в листе '{sheet}'")
               common_keys = set(df_main[key_col].dropna()) & set(
                   df_extra[key_col].dropna()
               )
               df_main = df_main[~df_main[key_col].isin(common_keys)]
               if sheet in VARIABLE_SHEETS:
                   target_cols = unified_columns[sheet]
                   aligned_extra = pd.DataFrame(
                       {
                           col: df_extra[col] if col in df_extra.columns else None
                           for col in target_cols
                       }
                   )
                   df_extra = aligned_extra
               result_data[sheet] = pd.concat([df_main, df_extra], ignore_index=True)
               current_step += 1
               self.update_progress(int(current_step / total_steps * 100))
           # Сохранение
           output_path = os.path.join(
               os.path.dirname(self.main_file), "merged_result.xlsx"
           )
           with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
               for sheet in SHEETS:
                   result_data[sheet].to_excel(
                       writer, sheet_name=sheet, startrow=1, index=False, header=True
                   )
           self.update_progress(100)
           self.status_label.config(text="Готово!")
           if missing_column_warnings:
               warn_msg = "Обнаружены различия в структуре таблиц:\n" + "\n".join(
                   missing_column_warnings
               )
               self.root.after(
                   0, lambda: messagebox.showwarning("Предупреждение", warn_msg)
               )
           self.root.after(0, lambda: self.show_completion(output_path))
       except Exception as e:
           self.root.after(
               0, lambda: messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")
           )
   def update_progress(self, value):
       self.root.after(0, lambda: self.progress.config(value=value))
       self.root.after(
           0, lambda: self.status_label.config(text=f"Выполнение: {value}%")
       )
   def show_completion(self, path):
       messagebox.showinfo("Готово", f"Склейка завершена!\nРезультат: {path}")
       self.root.destroy()
       self.parent_root.deiconify()

# --- Режим 3: Простая склейка импорта (с особым правилом для 'options') ---
class ExcelImportMergerApp:
   def __init__(self, root, parent_root):
       self.root = root
       self.parent_root = parent_root
       self.root.title("Склейка импорта (одинаковые файлы)")
       self.root.geometry("500x200")
       self.file_paths = []
       self.label = tk.Label(
           root, text="Выберите Excel-файлы для простой склейки", font=("Arial", 12)
       )
       self.label.pack(pady=20)
       self.select_button = tk.Button(
           root, text="Выбрать файлы", command=self.select_files, font=("Arial", 10)
       )
       self.select_button.pack(pady=5)
       self.progress = ttk.Progressbar(
           root, orient="horizontal", length=400, mode="determinate"
       )
       self.progress.pack(pady=10)
       self.status_label = tk.Label(root, text="", font=("Arial", 10))
       self.status_label.pack(pady=5)
   def select_files(self):
       self.file_paths = filedialog.askopenfilenames(
           title="Выберите Excel-файлы", filetypes=[("Excel files", "*.xlsx")]
       )
       if self.file_paths:
           self.select_button.config(state="disabled")
           self.label.config(text=f"Выбрано файлов: {len(self.file_paths)}")
           threading.Thread(target=self.merge_files, daemon=True).start()
       else:
           self.label.config(text="Файлы не выбраны")
   def merge_files(self):
       try:
           if not self.file_paths:
               raise Exception("Не выбрано ни одного файла")
           # Определяем листы по первому файлу
           with pd.ExcelFile(self.file_paths[0]) as first_xls:
               first_sheet_names = [name.strip() for name in first_xls.sheet_names]
           sheet_map = {name.strip().lower(): name for name in first_sheet_names}
           total_steps = len(self.file_paths) * len(sheet_map)
           current_step = 0
           merged_data = {name: [] for name in sheet_map.values()}
           for idx, file_path in enumerate(self.file_paths):
               try:
                   with pd.ExcelFile(file_path) as xls:
                       file_sheets = {
                           name.strip().lower(): name for name in xls.sheet_names
                       }
                       # Проверка: все листы из первого файла должны быть во всех остальных
                       for norm_name, orig_name in sheet_map.items():
                           if norm_name not in file_sheets:
                               raise ValueError(
                                   f"В файле {os.path.basename(file_path)} отсутствует лист '{orig_name}'"
                               )
                       for norm_name, orig_name in sheet_map.items():
                           # Особое правило: лист 'options' — только из первого файла
                           if orig_name.lower() == "options" and idx > 0:
                               current_step += 1
                               continue  # пропускаем
                           real_sheet_name = file_sheets[norm_name]
                           df = pd.read_excel(
                               xls, sheet_name=real_sheet_name, header=0
                           )
                           df.dropna(how="all", inplace=True)
                           merged_data[orig_name].append(df)
                           current_step += 1
                           self.update_progress(int(current_step / total_steps * 100))
               except Exception as e:
                   raise Exception(f"Ошибка при чтении {file_path}: {str(e)}")
           # Финальное объединение с учётом правила для 'options'
           final_dfs = {}
           for sheet_name, dfs in merged_data.items():
               if sheet_name.lower() == "options":
                   final_dfs[sheet_name] = dfs[0] if dfs else pd.DataFrame()
               else:
                   final_dfs[sheet_name] = (
                       pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
                   )
           # Сохранение
           output_path = os.path.join(
               os.path.dirname(self.file_paths[0]), "import_merged_result.xlsx"
           )
           with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
               for sheet_name, df in final_dfs.items():
                   df.to_excel(writer, sheet_name=sheet_name, index=False)
           self.update_progress(100)
           self.root.after(0, lambda: self.show_completion(output_path))
       except Exception as e:
           self.root.after(
               0, lambda: messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")
           )
   def update_progress(self, value):
       self.root.after(0, lambda: self.progress.config(value=value))
       self.root.after(
           0, lambda: self.status_label.config(text=f"Обработка: {value}%")
       )
   def show_completion(self, path):
       messagebox.showinfo("Готово", f"Файлы успешно склеены!\nРезультат: {path}")
       self.root.destroy()
       self.parent_root.deiconify()

def main():
   root = tk.Tk()
   app = MainApp(root)
   root.mainloop()

if __name__ == "__main__":
   main()
