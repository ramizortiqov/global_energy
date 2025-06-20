import customtkinter as ctk
import psycopg2
from tkinter import messagebox, ttk, filedialog
import os
from dotenv import load_dotenv
import csv # Import the csv module

# Импорты для анализа и визуализации
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import pandas as pd
import seaborn as sns

# Загрузка переменных окружения из .env файла
load_dotenv()

# --- Константы и конфигурация ---
METRIC_COLUMN_MAP = {
    "total_energy_twh": "Общее потребление (TWh)",
    "energy_per_capita_kwh": "Потребление на душу (kWh)",
    "renewable_share_pct": "Доля ВИЭ (%)",
    "fossil_fuel_pct": "Зависимость от топлива (%)",
    "industry_pct": "Доля индустрии (%)",
    "household_pct": "Доля хозяйств (%)",
    "co2_emissions_mt": "Выбросы CO2 (Мт)",
    "energy_price_usd": "Цена энергии (USD/kWh)"
}
# Создаем список ключей для сохранения порядка
METRIC_KEYS = list(METRIC_COLUMN_MAP.keys())

# --- Функции для работы с базой данных (остаются без изменений) ---
def get_db_connection():
    try:
        conn = psycopg2.connect(
            dbname=os.getenv("DB_NAME"), user=os.getenv("DB_USER"),
            password=os.getenv("DB_PASSWORD"), host=os.getenv("DB_HOST"),
            port=os.getenv("DB_PORT")
        )
        return conn
    except psycopg2.OperationalError as e:
        messagebox.showerror("Ошибка подключения", f"Не удалось подключиться к базе данных:\n{e}")
        return None

def fetch_countries():
    conn = get_db_connection()
    if conn is None: return []
    countries = []
    try:
        cur = conn.cursor()
        cur.execute("SELECT country_id, country_name FROM Countries ORDER BY country_name")
        countries = cur.fetchall()
        cur.close()
    except psycopg2.Error as e: messagebox.showerror("Ошибка запроса", f"Не удалось получить список стран:\n{e}")
    finally:
        if conn: conn.close()
    return countries

def fetch_metrics_for_country(country_id):
    conn = get_db_connection()
    if conn is None: return []
    metrics = []
    try:
        cur = conn.cursor()
        sql = "SELECT metric_id, year, total_energy_twh, energy_per_capita_kwh, renewable_share_pct, fossil_fuel_pct, industry_pct, household_pct, co2_emissions_mt, energy_price_usd FROM Energy_Metrics WHERE country_id = %s ORDER BY year ASC"
        cur.execute(sql, (country_id,))
        metrics = cur.fetchall()
        cur.close()
    except psycopg2.Error as e: messagebox.showerror("Ошибка запроса", f"Не удалось получить данные для страны:\n{e}")
    finally:
        if conn: conn.close()
    return metrics

def fetch_distinct_years():
    conn = get_db_connection()
    if conn is None: return []
    years = []
    try:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT year FROM Energy_Metrics ORDER BY year DESC")
        years = [row[0] for row in cur.fetchall()]
        cur.close()
    except psycopg2.Error as e: messagebox.showerror("Ошибка запроса", f"Не удалось получить список лет:\n{e}")
    finally:
        if conn: conn.close()
    return years

def fetch_yearly_ranking(year, order_metric_column):
    conn = get_db_connection()
    if conn is None: return []
    ranking_data = []
    try:
        cur = conn.cursor()
        if order_metric_column not in METRIC_COLUMN_MAP.keys(): raise ValueError("Недопустимая колонка для сортировки")
        all_metric_columns = ", ".join([f"m.{col}" for col in METRIC_KEYS])
        sql = f"SELECT c.country_name, m.year, {all_metric_columns} FROM Energy_Metrics m JOIN Countries c ON m.country_id = c.country_id WHERE m.year = %s ORDER BY m.{order_metric_column} DESC"
        cur.execute(sql, (year,))
        ranking_data = cur.fetchall()
        cur.close()
    except (psycopg2.Error, ValueError) as e: messagebox.showerror("Ошибка запроса", f"Не удалось получить рейтинг:\n{e}")
    finally:
        if conn: conn.close()
    return ranking_data

def execute_query(sql, params=None, message="Запрос выполнен успешно!"):
    conn = get_db_connection()
    if conn is None: return False
    success = False
    try:
        cur = conn.cursor()
        cur.execute(sql, params)
        conn.commit()
        cur.close()
        messagebox.showinfo("Успех", message)
        success = True
    except psycopg2.Error as e: messagebox.showerror("Ошибка запроса", f"Не удалось выполнить запрос:\n{e}")
    finally:
        if conn: conn.close()
    return success

# --- Окно редактирования/добавления ---
class EditWindow(ctk.CTkToplevel):
    def __init__(self, master, country_id, record=None):
        super().__init__(master)
        self.master_app = master; self.country_id = country_id; self.record = record
        self.title("Редактирование записи" if record else "Добавление записи")
        self.geometry("450x550")
        self.entries = {}
        self.fields_config = [("Год", "year")] + list(METRIC_COLUMN_MAP.items())
        for i, (text, key) in enumerate(self.fields_config):
            label = ctk.CTkLabel(self, text=text); label.grid(row=i, column=0, padx=10, pady=10, sticky="w")
            entry = ctk.CTkEntry(self); entry.grid(row=i, column=1, padx=10, pady=10, sticky="ew")
            self.entries[key] = entry
        if self.record:
            for i, (_, key) in enumerate(self.fields_config): self.entries[key].insert(0, self.record[i+1])
        save_button = ctk.CTkButton(self, text="Сохранить", command=self.save); save_button.grid(row=len(self.fields_config), column=0, columnspan=2, pady=20)
        self.grab_set()

    def save(self):
        data_to_save = {}
        try:
            for text, key in self.fields_config:
                value_str = self.entries[key].get().replace(',', '.')
                if not value_str: messagebox.showerror("Ошибка ввода", f"Поле '{text}' не может быть пустым."); return
                data_to_save[key] = int(value_str) if key == 'year' else float(value_str)
        except ValueError: messagebox.showerror("Ошибка ввода", "Пожалуйста, убедитесь, что все поля содержат корректные числовые значения."); return
        values = tuple(data_to_save[key] for _, key in self.fields_config)
        
        db_cols_str = ", ".join(["year"] + METRIC_KEYS)
        set_str = ", ".join([f"{key}=%s" for key in ["year"] + METRIC_KEYS])
        placeholders_str = ", ".join(["%s"] * (len(METRIC_KEYS) + 2))
        
        if self.record:
            metric_id = self.record[0]
            sql = f"UPDATE Energy_Metrics SET {set_str} WHERE metric_id = %s"
            if execute_query(sql, values + (metric_id,), "Запись успешно обновлена!"): self.master_app.refresh_data(); self.destroy()
        else:
            sql = f"INSERT INTO Energy_Metrics (country_id, {db_cols_str}) VALUES ({placeholders_str})"
            if execute_query(sql, (self.country_id,) + values, "Запись успешно добавлена!"): self.master_app.refresh_data(); self.destroy()

# --- Основной класс приложения ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Аналитическая панель 'Глобальное энергопотребление'")
        self.geometry("1600x950")
        
        # --- !!! НОВЫЙ БЛОК: НАСТРОЙКА СТИЛЕЙ ТАБЛИЦЫ !!! ---
        style = ttk.Style()
        
        # Выбираем тему, которую CustomTkinter использует по умолчанию, чтобы наши стили ее дополняли
        # На Windows это обычно 'vista', на Mac 'aqua', на Linux 'clam'
        # Мы можем установить 'clam' для кросс-платформенной консистентности
        style.theme_use("clam") 
        # Поместите этот код в метод __init__ класса App

        # --- НАСТРОЙКА СТИЛЕЙ WIDGETS ---
        style = ttk.Style()

        # Используем тему 'clam', так как она лучше всего поддается настройке
        # и хорошо смотрится в связке с CustomTkinter.
        style.theme_use("clam")

        # --- НАСТРОЙКА СТИЛЯ ЗАГОЛОВКОВ ТАБЛИЦЫ ---
        style.configure("Treeview.Heading",
                        font=('Segoe UI', 11, 'bold'),  # Шрифт: Segoe UI, размер 11, жирный
                        background="#3a7ebf",           # Приятный синий цвет фона
                        foreground="white",              # Белый цвет текста
                        relief="flat")                   # Плоский вид, без объемных границ

        # --- (БОНУС) Изменение цвета при наведении мыши ---
        # Эта настройка делает заголовок чуть темнее, когда на него наведен курсор
        style.map("Treeview.Heading",
                background=[('active', '#346c9f')])
        # Настраиваем стиль заголовков
        style.configure("Treeview.Heading",
                        font=('Verdana', 11, 'bold'),
                        background="#2c5c8c",   # Темно-синий фон
                        foreground="white")   # Белый текст

        # # Убираем стандартные рамки заголовков в теме 'clam'
        # style.layout("Treeview.Heading", [
        #     ('Treeview.heading.cell', {'sticky': 'nswe'})
        # ])
        
        # Настраиваем стиль основной таблицы
        style.configure("Treeview",
                        rowheight=28,   # Увеличиваем высоту строк
                        font=('Verdana', 10),
                        fieldbackground="#333333", # Фон ячеек в темной теме
                        background="#2b2b2b",       # Общий фон таблицы
                        foreground="white")       # Цвет текста

        # Создаем тег для нечетных строк, чтобы сделать их другого цвета
        style.configure("oddrow.TTreeview",
                        background="#3a3d3e") # Слегка другой оттенок серого

        # Применяем стиль для выделенной строки
        style.map('Treeview',
                background=[('selected', '#2c5c8c')])
        
        self.current_country_id = None; self.current_country_name = None; self.current_metrics_data = []

        # --- ОСНОВНАЯ СТРУКТУРА ОКНА ---
        self.grid_rowconfigure(3, weight=1); self.grid_columnconfigure(0, weight=1)
        self.top_frame = ctk.CTkFrame(self); self.top_frame.grid(row=0, column=0, pady=10, padx=20, sticky="ew")
        self.ranking_frame = ctk.CTkFrame(self); self.ranking_frame.grid(row=1, column=0, pady=(0, 10), padx=20, sticky="ew")
        self.stats_frame = ctk.CTkScrollableFrame(self, label_text="Сводная статистика"); self.stats_frame.grid(row=2, column=0, pady=(0, 10), padx=20, sticky="ew")
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent"); self.main_frame.grid(row=3, column=0, pady=10, padx=20, sticky="nsew")
        self.crud_frame = ctk.CTkFrame(self); self.crud_frame.grid(row=4, column=0, pady=10, padx=20, sticky="ew")

        # --- ЗАПОЛНЕНИЕ ПАНЕЛЕЙ ---
        ctk.CTkLabel(self.top_frame, text="Анализ по стране:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 10))
        self.country_data = fetch_countries()
        country_names = [c[1] for c in self.country_data] if self.country_data else ["Нет данных"]
        self.country_menu = ctk.CTkOptionMenu(self.top_frame, values=country_names, command=self.on_country_select)
        self.country_menu.set("Выберите страну..."); self.country_menu.pack(side="left")

        ctk.CTkLabel(self.ranking_frame, text="Рейтинг по году:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 10))
        self.years_data = fetch_distinct_years()
        self.year_menu = ctk.CTkOptionMenu(self.ranking_frame, values=[str(y) for y in self.years_data]); self.year_menu.pack(side="left", padx=5)
        self.metric_menu = ctk.CTkOptionMenu(self.ranking_frame, values=list(METRIC_COLUMN_MAP.values())); self.metric_menu.pack(side="left", padx=5)
        ctk.CTkButton(self.ranking_frame, text="Показать рейтинг", command=self.show_ranking).pack(side="left", padx=5)

        # !!! НОВОЕ: Динамическое создание панели статистики
        self.stat_labels = {}
        for key, text in METRIC_COLUMN_MAP.items():
            frame = ctk.CTkFrame(self.stats_frame)
            frame.pack(fill="x", padx=5, pady=2)
            ctk.CTkLabel(frame, text=text, anchor="w", width=250).pack(side="left")
            self.stat_labels[key] = {
                "avg": ctk.CTkLabel(frame, text="Среднее: --", anchor="w", width=200),
                "min": ctk.CTkLabel(frame, text="Мин: --", anchor="w", width=200),
                "max": ctk.CTkLabel(frame, text="Макс: --", anchor="w", width=200)
            }
            self.stat_labels[key]["avg"].pack(side="left", padx=5)
            self.stat_labels[key]["min"].pack(side="left", padx=5)
            self.stat_labels[key]["max"].pack(side="left", padx=5)

        # ОСНОВНАЯ РАБОЧАЯ ОБЛАСТЬ (ВКЛАДКИ)
        self.tab_view = ctk.CTkTabview(self.main_frame); self.tab_view.pack(fill="both", expand=True)
        self.tab_view.add("Таблица"); self.tab_view.add("График")
        
        self.table_container_frame = ctk.CTkFrame(self.tab_view.tab("Таблица"), fg_color="transparent"); self.table_container_frame.pack(fill="both", expand=True)
        self.search_entry = ctk.CTkEntry(self.table_container_frame, placeholder_text="Фильтр по данным..."); self.search_entry.pack(fill="x", padx=5, pady=5); self.search_entry.bind("<KeyRelease>", self.filter_data)
        self.tree = ttk.Treeview(self.table_container_frame, show="headings"); self.tree.pack(fill="both", expand=True, padx=5, pady=5)
        
        # !!! НОВОЕ: Интерактивный график
        self.plot_frame = ctk.CTkFrame(self.tab_view.tab("График"), fg_color="transparent"); self.plot_frame.pack(fill="both", expand=True)
        self.plot_controls_frame = ctk.CTkFrame(self.plot_frame); self.plot_controls_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(self.plot_controls_frame, text="Показать на графике:").pack(side="left", padx=10)
        self.plot_metric_menu = ctk.CTkOptionMenu(self.plot_controls_frame, values=list(METRIC_COLUMN_MAP.values()), command=lambda _: self.plot_data()); self.plot_metric_menu.pack(side="left")
        
        self.fig = Figure(figsize=(5, 4), dpi=100); self.ax = self.fig.add_subplot(111); self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_frame); self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self.reconfigure_tree_for_country_view()
        
        # НИЖНЯЯ ПАНЕЛЬ (КНОПКИ)
        ctk.CTkButton(self.crud_frame, text="Добавить", command=self.add_record).pack(side="left", expand=True, padx=5, pady=5)
        ctk.CTkButton(self.crud_frame, text="Изменить", command=self.edit_record).pack(side="left", expand=True, padx=5, pady=5)
        ctk.CTkButton(self.crud_frame, text="Удалить", command=self.delete_record).pack(side="left", expand=True, padx=5, pady=5)
        ctk.CTkButton(self.crud_frame, text="Экспорт в CSV", command=self.export_data_to_csv).pack(side="left", expand=True, padx=5, pady=5) # New button
        # ctk.CTkButton(self.crud_frame, text="Показать корреляцию", command=self.show_correlation).pack(side="left", expand=True, padx=5, pady=5)
    
    # --- МЕТОДЫ ПРИЛОЖЕНИЯ ---
    def reconfigure_tree_for_country_view(self):
        columns = ("year",) + tuple(METRIC_KEYS); self.tree.configure(columns=columns)
        self.tree.heading("year", text="Год"); self.tree.column("year", width=60, anchor="center")
        for col_key, col_text in METRIC_COLUMN_MAP.items(): self.tree.heading(col_key, text=col_text); self.tree.column(col_key, width=150, anchor="center")
    
    def reconfigure_tree_for_ranking_view(self):
        columns = ("country", "year") + tuple(METRIC_KEYS); self.tree.configure(columns=columns)
        self.tree.heading("country", text="Страна"); self.tree.column("country", width=120)
        self.tree.heading("year", text="Год"); self.tree.column("year", width=60)
        for col_key, col_text in METRIC_COLUMN_MAP.items(): self.tree.heading(col_key, text=col_text); self.tree.column(col_key, width=150, anchor="center")

    # В классе App
    def populate_treeview(self, data, mode='country'):
        # Очищаем таблицу от старых данных
        for i in self.tree.get_children():
            self.tree.delete(i)

        # Вставляем новые данные с тегами для чередования цветов
        if mode == 'country':
            for i, row in enumerate(data):
                # Применяем тег 'oddrow' к каждой нечетной строке (i % 2 != 0)
                tag = 'oddrow' if i % 2 != 0 else ''
                self.tree.insert("", "end", iid=row[0], values=row[1:], tags=(tag,))
        elif mode == 'ranking':
            for i, row in enumerate(data):
                tag = 'oddrow' if i % 2 != 0 else ''
                self.tree.insert("", "end", values=row, tags=(tag,))

    def filter_data(self, event=None):
        search_term = self.search_entry.get().lower()
        if not search_term: self.populate_treeview(self.current_metrics_data)
        else:
            filtered_data = [row for row in self.current_metrics_data if any(search_term in str(val).lower() for val in row)]
            self.populate_treeview(filtered_data)

    def on_country_select(self, name):
        self.reconfigure_tree_for_country_view()
        self.current_country_name = name
        self.current_country_id = next((cid for cid, cname in self.country_data if cname == name), None)
        self.refresh_data()
        
    def refresh_data(self):
        self.search_entry.delete(0, "end")
        if self.current_country_id is None: self.current_metrics_data = []; self.populate_treeview([]); self.plot_data(); self.update_statistics([]); return
        self.current_metrics_data = fetch_metrics_for_country(self.current_country_id)
        self.populate_treeview(self.current_metrics_data)
        self.plot_data()
        self.update_statistics()
    
    def show_ranking(self):
        self.reconfigure_tree_for_ranking_view(); self.clear_country_specific_views()
        year = self.year_menu.get()
        metric_friendly_name = self.metric_menu.get()
        metric_column = next((key for key, value in METRIC_COLUMN_MAP.items() if value == metric_friendly_name), None)
        if not metric_column: messagebox.showerror("Ошибка", "Не удалось определить поле для сортировки."); return
        ranking_data = fetch_yearly_ranking(year, metric_column)
        self.populate_treeview(ranking_data, mode='ranking')

    # !!! ИЗМЕНЕНИЕ: Логика графика стала полностью динамической
    def plot_data(self):
        self.ax.clear()
        if self.current_metrics_data:
            metric_friendly_name = self.plot_metric_menu.get()
            metric_key = next((key for key, value in METRIC_COLUMN_MAP.items() if value == metric_friendly_name), None)
            if metric_key:
                # Находим индекс нужной колонки. +2 потому что в self.current_metrics_data есть metric_id и year
                metric_index = METRIC_KEYS.index(metric_key) + 2 
                years = [row[1] for row in self.current_metrics_data]
                values = [row[metric_index] for row in self.current_metrics_data]
                self.ax.plot(years, values, marker='o', linestyle='-')
                self.ax.set_title(f"{metric_friendly_name} для {self.current_country_name}")
                self.ax.set_xlabel("Год")
                self.ax.set_ylabel(metric_friendly_name)
                self.ax.grid(True)
        self.fig.tight_layout()
        self.canvas.draw()
        
    # !!! ИЗМЕНЕНИЕ: Логика статистики стала полностью динамической
    def update_statistics(self):
        if not self.current_metrics_data:
            for key in self.stat_labels:
                self.stat_labels[key]["avg"].configure(text="Среднее: --")
                self.stat_labels[key]["min"].configure(text="Мин: --")
                self.stat_labels[key]["max"].configure(text="Макс: --")
            return
        
        for i, key in enumerate(METRIC_KEYS):
            metric_index = i + 2 # +2 из-за metric_id и year
            column_data = [row[metric_index] for row in self.current_metrics_data]
            
            avg_val = np.mean(column_data)
            min_val = min(column_data)
            max_val = max(column_data)
            
            min_year = self.current_metrics_data[column_data.index(min_val)][1]
            max_year = self.current_metrics_data[column_data.index(max_val)][1]

            self.stat_labels[key]["avg"].configure(text=f"Среднее: {avg_val:.2f}")
            self.stat_labels[key]["min"].configure(text=f"Мин: {min_val:.2f} ({min_year})")
            self.stat_labels[key]["max"].configure(text=f"Макс: {max_val:.2f} ({max_year})")

    def clear_country_specific_views(self):
        self.current_country_id = None; self.current_country_name = "Рейтинг"; self.current_metrics_data = []
        for i in self.tree.get_children(): self.tree.delete(i)
        self.plot_data(); self.update_statistics()
        self.country_menu.set("Выберите страну...")
        
    # def show_correlation(self):
    #     if not self.current_metrics_data: messagebox.showwarning("Внимание", "Сначала выберите страну и загрузите данные."); return
    #     columns = ["id", "year"] + METRIC_KEYS
    #     df = pd.DataFrame(self.current_metrics_data, columns=columns).drop(columns=['id', 'year'])
    #     corr_window = ctk.CTkToplevel(self); corr_window.title(f"Матрица корреляций для {self.current_country_name}"); corr_window.geometry("800x600")
    #     fig = Figure(figsize=(8, 6), dpi=100); ax = fig.add_subplot(111)
    #     sns.heatmap(df.corr(), annot=True, cmap='viridis', fmt=".2f", ax=ax); fig.tight_layout()
    #     canvas = FigureCanvasTkAgg(fig, master=corr_window); canvas.draw(); canvas.get_tk_widget().pack(fill="both", expand=True)
    #     corr_window.grab_set()

    # --- Методы для CRUD ---
    def add_record(self):
        if self.current_country_id is None: messagebox.showwarning("Внимание", "Сначала выберите страну для добавления записи."); return
        EditWindow(self, self.current_country_id)
    def edit_record(self):
        selected_iid = self.tree.focus();
        if not selected_iid: messagebox.showwarning("Внимание", "Выберите запись для редактирования."); return
        if self.current_country_name == "Рейтинг": messagebox.showwarning("Внимание", "Редактирование недоступно в режиме рейтинга."); return
        full_record = next((row for row in self.current_metrics_data if row[0] == int(selected_iid)), None)
        if full_record: EditWindow(self, self.current_country_id, record=full_record)
    def delete_record(self):
        selected_iid = self.tree.focus();
        if not selected_iid: messagebox.showwarning("Внимание", "Выберите запись для удаления."); return
        if self.current_country_name == "Рейтинг": messagebox.showwarning("Внимание", "Удаление недоступно в режиме рейтинга."); return
        if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить эту запись?"):
            if execute_query("DELETE FROM Energy_Metrics WHERE metric_id = %s", (selected_iid,), "Запись успешно удалена!"): self.refresh_data()

    def export_data_to_csv(self):
        """
        Exports the data currently displayed in the Treeview to a CSV file.
        """
        # Get the columns from the Treeview
        columns = [self.tree.heading(col, "text") for col in self.tree["columns"]]
        
        # Get the data from the Treeview
        data = []
        for item_id in self.tree.get_children():
            data.append(self.tree.item(item_id, "values"))

        if not data:
            messagebox.showwarning("Внимание", "Нет данных для экспорта.");
            return

        # Ask the user for a file name and location
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not file_path: # User cancelled the dialog
            return

        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(columns) # Write headers
                writer.writerows(data)   # Write data rows
            messagebox.showinfo("Успех", f"Данные успешно экспортированы в {file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", f"Не удалось экспортировать данные:\n{e}")


# --- Запуск приложения ---
if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("dark-blue")
    app = App()
    app.mainloop()