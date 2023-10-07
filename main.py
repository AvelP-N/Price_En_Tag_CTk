import os
import re
import tkinter
import customtkinter
import xml.etree.ElementTree as ET
from xlrd import open_workbook
from transliterate import translit


class XlsBook:
    """Класс для сбора данных из XLS файла"""

    workbook = None
    list_sheets = []
    sheet = None
    rows = None
    cols = None
    code = None
    name = None
    price = None
    term = None
    count_tag = 0

    def load_workbook(self, file_path):
        """Загрузка объекта книги и список листов"""

        self.workbook = open_workbook(file_path)
        self.list_sheets = self.workbook.sheet_names()
        return self.list_sheets

    def get_sheet_rows_cols(self, sheet_name):
        """Получить объект листа, колонок и столбцов"""

        self.sheet = self.workbook.sheet_by_name(sheet_name)
        self.rows = self.sheet.nrows
        self.cols = self.sheet.ncols

    def check_selected_tag(self):
        """Проверка, какие теги выбраны и их количество"""

        tags = self.get_count_tag(self.code, self.name, self.price, self.term)

        if tags < 3:
            return "Count tags < 3"
        elif tags == 3:
            if self.get_count_tag(self.code, self.name, self.price) == 3:
                self.count_tag = 3
            else:
                return "Tags are required: < testShortName >, < testName >, < testPrice >"
        else:
            self.count_tag = 4

    @staticmethod
    def get_count_tag(*args):
        """Получить количество тегов для создания XML файла"""

        tags = 0

        for tag in args:
            if isinstance(tag, int):
                tags += 1

        return tags


class XmlFile(XlsBook):
    """Класс для создания XML файла"""

    root = ET.Element('root')
    root.set('xmlns:xsi', "http://www.w3.org/2001/XMLSchema-instance")
    found_code_test = 0
    correction_data = ""

    def create_xml_tree(self):
        """Метод для создания XML файла"""

        self.correction_data = ""

        for n_row in range(self.rows):
            row = self.sheet.row_slice(n_row, 0, self.cols)
            row_code = str(row[self.code].value)
            if row_code.count('.') >= 2 and len(row_code) < 15:
                self.found_code_test += 1

                main_price = ET.SubElement(self.root, "price")

                # Создать тег testShortName и выполнить проверку на русские буквы
                short_name = ET.SubElement(main_price, "testShortName")
                short_name.text = self.check_code_ru(str(row[self.code].value).strip())

                # Создать тег testName
                name = ET.SubElement(main_price, "testName")
                name.text = str(row[self.name].value)

                # Создать тег testPrice и проверка цены
                price = ET.SubElement(main_price, "testPrice")
                price.text = self.check_price(str(row[self.price].value).strip(), str(row[self.code].value).strip())

                # Создать тег term
                if isinstance(self.term, int):
                    term = ET.SubElement(main_price, 'term')
                    term.text = self.check_deadline(row[self.term], row_code)

    def check_code_ru(self, data):
        """Проверка кода теста на русские буквы и замена на латинские"""

        if data.isascii():
            return data

        edit_code = translit(data.upper(), reversed=True)
        self.correction_data += f'{data}  -  < Ru letters >  changed  < {edit_code} >\n'

        return edit_code

    def check_price(self, data_price, data_short_name):
        """Редактирование цены. Убрать копейки и пробелы в рублях"""

        split_price = re.split('[.,]', data_price)[0]
        edit_price = ''.join(re.findall(r'\d', split_price))

        if edit_price.isdigit():
            return edit_price

        self.correction_data += f"{data_short_name}  -  < {(data_price, 'Empty price')[data_price == '']} >  " \
                                f"changed  < 0 >\n"
        return '0'

    def check_deadline(self, data, code_test):
        """Проверка сроков выполнения тестов. Если нет срока, то подставить ноль"""

        if data.ctype == 1:  # Если тип данных строка
            edit_deadline = re.findall(r'\d+', data.value)
            if edit_deadline:
                self.correction_data += f"{code_test}  -  Deadline < {data.value} > changed < {edit_deadline[-1]}\n"
                return edit_deadline[-1]
        elif data.ctype == 2:  # Если тип данных числовой
            return str(int(float(data.value)))

        self.correction_data += f"{code_test}  -  < Empty deadline > changed  < 0 >\n"

        return '0'


class App(customtkinter.CTk, XmlFile, XlsBook,):
    """Класс для отрисовки программы"""

    def __init__(self):
        super().__init__()

        self.title("Price En Tag v1.0")
        self.resizable(False, False)
        customtkinter.set_appearance_mode("dark")

        self.top_label = customtkinter.CTkLabel(self, text="Create an XML file with English tags",
                                                font=("Times New Roman", 22, "bold"), text_color="gray")
        self.top_label.pack(padx=10, pady=10)

        # Рамка, где есть кнопка открыть файл и выпадающий список с выбором листа
        self.top_frame = customtkinter.CTkFrame(self)
        self.top_frame.pack(padx=10, pady=10, fill=tkinter.BOTH)

        self.button = customtkinter.CTkButton(master=self.top_frame, text="Open file", command=self.open_file,
                                              text_color="black", font=("Times New Roman", 14, "bold"))
        self.button.grid(row=0, column=0, padx=10, pady=10)

        self.file_label = customtkinter.CTkLabel(master=self.top_frame, text="Select the *.xls file")
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        select_sheet_var = tkinter.StringVar(value="Select sheet")
        self.select_sheet = customtkinter.CTkComboBox(master=self.top_frame, state="disabled",
                                                      variable=select_sheet_var, justify="center",
                                                      command=self.get_sheet_cols)
        self.select_sheet.grid(row=1, column=0, padx=10, pady=10)

        self.label_sheet = customtkinter.CTkLabel(master=self.top_frame, text="")
        self.label_sheet.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # Рамка с выбором столбцов
        self.middle_frame = customtkinter.CTkFrame(self)
        self.middle_frame.pack(padx=10, pady=10, fill=tkinter.BOTH)

        self.test_short_name_label = customtkinter.CTkLabel(master=self.middle_frame, width=90, text="testShortName",
                                                            text_color="grey", font=("Times New Roman", 14, "bold"))
        self.test_short_name_label.grid(row=0, column=0, padx=10)

        self.test_name_label = customtkinter.CTkLabel(master=self.middle_frame, width=90, text="testName",
                                                      text_color="grey", font=("Times New Roman", 14, "bold"))
        self.test_name_label.grid(row=0, column=1, padx=10)

        self.test_price_label = customtkinter.CTkLabel(master=self.middle_frame, width=90, text="testPrice",
                                                       text_color="grey", font=("Times New Roman", 14, "bold"))
        self.test_price_label.grid(row=0, column=2, padx=10)

        self.test_term_label = customtkinter.CTkLabel(master=self.middle_frame, width=90, text="term",
                                                      text_color="grey", font=("Times New Roman", 14, "bold"))
        self.test_term_label.grid(row=0, column=3, padx=10)

        self.test_short_name_box = customtkinter.CTkComboBox(master=self.middle_frame, values=["0"], width=60,
                                                             text_color="black", command=self.get_test_short_name)
        self.test_short_name_box.grid(row=1, column=0, padx=10, pady=10)

        self.test_name_box = customtkinter.CTkComboBox(master=self.middle_frame, values=["0"], width=60,
                                                       text_color="black", command=self.get_test_name)
        self.test_name_box.grid(row=1, column=1, padx=10, pady=10)

        self.test_price_box = customtkinter.CTkComboBox(master=self.middle_frame, values=["0"], width=60,
                                                        text_color="black", command=self.get_test_price)
        self.test_price_box.grid(row=1, column=2, padx=10, pady=10)

        self.test_term_box = customtkinter.CTkComboBox(master=self.middle_frame, values=["0"], width=60,
                                                       text_color="black", command=self.get_test_term)
        self.test_term_box.grid(row=1, column=3, padx=10, pady=10)

        # Рамка с кнопками парсинга XLS листов и создания XML файла
        self.bottom_frame = customtkinter.CTkFrame(self)
        self.bottom_frame.pack(padx=10, pady=10, fill=tkinter.BOTH)

        self.button_pars = customtkinter.CTkButton(master=self.bottom_frame, text="Pars sheet", text_color="black",
                                                   font=("Times New Roman", 14, "bold"), command=self.button_pars_sheet)
        self.button_pars.pack(side="left", padx=10, pady=10)

        self.button_create_xml = customtkinter.CTkButton(master=self.bottom_frame, text="Create XML", text_color="black",
                                                         font=("Times New Roman", 14, "bold"),
                                                         command=self.create_xml_file)
        self.button_create_xml.pack(side="right", padx=10, pady=10)

        # Нижняя текстовая рамка, для вывода информации
        self.text_box = customtkinter.CTkTextbox(self, height=150, text_color="green")
        self.text_box.pack(padx=10, pady=(0, 10), fill=tkinter.BOTH)

    def open_file(self):
        """Получить путь до файла XLS книги и разблокировать CTkComboBox во второй строке"""

        file_path = customtkinter.filedialog.askopenfilename()

        if file_path:
            # Сбросить значения на дефолтные при открытии нового файла
            self.default_params()

            if os.path.splitext(file_path)[1].lower() == ".xls":
                file = f"File selected:  {os.path.split(file_path)[1]}"
                self.file_label.configure(text=file, text_color="white")
                os.startfile(file_path)

                # Передать путь до файла и вернуть список листов
                sheets = self.load_workbook(file_path)

                self.select_sheet.configure(state="normal", values=sheets, text_color="white")
            else:
                self.file_label.configure(text="Please select the XLS file", text_color="red")
        else:
            self.text_box.insert(tkinter.END, f"File not selected!\n\n")
            self.text_box.see(tkinter.END)

    def default_params(self):
        """Установить значения по умолчанию, выбор листа и выбор колонок, когда открываем новый файл"""

        self.workbook = None
        self.list_sheets = []
        self.sheet = None
        self.rows = None
        self.cols = None
        self.code = None
        self.name = None
        self.price = None
        self.term = None
        self.count_tag = 0
        self.root = ET.Element('root')
        self.root.set('xmlns:xsi', "http://www.w3.org/2001/XMLSchema-instance")
        self.count_tag = 0
        self.found_code_test = 0
        self.correction_data = ""

        select_sheet_var = tkinter.StringVar(value="Select sheet")
        self.select_sheet.configure(variable=select_sheet_var)
        self.label_sheet.configure(text="")

        self.test_short_name_box.configure(values=["0"], text_color="black")
        self.test_short_name_box.set("0")
        self.test_name_box.configure(values=["0"], text_color="black")
        self.test_name_box.set("0")
        self.test_price_box.configure(values=["0"], text_color="black")
        self.test_price_box.set("0")
        self.test_term_box.configure(values=["0"], text_color="black")
        self.test_term_box.set("0")

        self.text_box.delete("0.0", "end")

    def get_sheet_cols(self, sheet):
        """Пользовательский выбор листа и колонок"""

        self.label_sheet.configure(text=f"Selected sheet:  {sheet}")

        # По названию листа получить объект листа количество строк и колонок
        self.get_sheet_rows_cols(sheet)

        self.text_box.insert(tkinter.END, f"Sheet - {self.sheet.name}\nRows - {self.rows}\nCols - {self.cols}\n\n")
        self.text_box.see(tkinter.END)

        # Установить в ComboBox список с колонками в выбранном листе и сбросить все значения на дефолтные
        list_column = list(map(str, range(self.cols + 1)))
        self.test_short_name_box.configure(values=list_column, text_color="black")
        self.test_short_name_box.set("0")
        self.test_name_box.configure(values=list_column, text_color="black")
        self.test_name_box.set("0")
        self.test_price_box.configure(values=list_column, text_color="black")
        self.test_price_box.set("0")
        self.test_term_box.configure(values=list_column, text_color="black")
        self.test_term_box.set("0")

    def get_test_short_name(self, number):

        if int(number) > 0:
            self.code = int(number) - 1
            self.test_short_name_box.configure(text_color="white")
        else:
            self.code = None
            self.test_short_name_box.configure(text_color="black")

    def get_test_name(self, number):

        if int(number) > 0:
            self.name = int(number) - 1
            self.test_name_box.configure(text_color="white")
        else:
            self.name = None
            self.test_name_box.configure(text_color="black")

    def get_test_price(self, number):

        if int(number) > 0:
            self.price = int(number) - 1
            self.test_price_box.configure(text_color="white")
        else:
            self.price = None
            self.test_price_box.configure(text_color="black")

    def get_test_term(self, number):

        if int(number) > 0:
            self.term = int(number) - 1
            self.test_term_box.configure(text_color="white")
        else:
            self.term = None
            self.test_term_box.configure(text_color="black")

    def button_pars_sheet(self):
        """Парсинг XLS листа и создание XML дерева"""

        # Проверить сколько тегов пользователь выбрал
        count_tags = self.check_selected_tag()

        if count_tags == "Count tags < 3":
            self.text_box.insert(tkinter.END, f"Count tags < 3\n")
            self.text_box.insert(tkinter.END, "Please select 3 or 4 tags!\n\n")
            self.text_box.see(tkinter.END)
        elif count_tags == "Tags are required: < testShortName >, < testName >, < testPrice >":
            self.text_box.insert(tkinter.END, f"{count_tags}\n\n")
            self.text_box.see(tkinter.END)
        else:
            self.create_xml_tree()

            self.text_box.insert(tkinter.END, f"{self.correction_data}\n")
            self.text_box.see(tkinter.END)

            self.text_box.insert(tkinter.END, f"Parser sheet < {self.sheet.name} > complete!\n\n")
            self.text_box.see(tkinter.END)

    def create_xml_file(self):
        """Создать XML файл"""

        tree = ET.ElementTree(self.root)
        tree.write(fr'C:\Users\{os.getlogin()}\Downloads\price.xml', encoding="UTF-8", xml_declaration=True)
        os.startfile(fr'C:\Users\{os.getlogin()}\Downloads\price.xml')

        self.text_box.insert(tkinter.END, f"Found CodeTest  -  {self.found_code_test}\n")
        self.text_box.insert(tkinter.END, f"Path XML file  -  C:\\Users\\{os.getlogin()}\\Downloads\\price.xml\n\n")

        self.text_box.insert(tkinter.END, "Текст для закрытия заявки:\n")
        self.text_box.insert(tkinter.END, "Прайс выгружен, сообщите интегратору о необходимости обновления "
                                          "в МИС клиента.\n\n")
        self.text_box.see(tkinter.END)


if __name__ == "__main__":
    app = App()
    app.mainloop()
