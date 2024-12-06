import sys
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QTextCursor, QTextCharFormat, QTextBlockFormat, QImage, QPixmap, QTextImageFormat, qRgba
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTextEdit
from PyQt5.QtPrintSupport import QPrinter
import os

os.environ['QT_PLUGIN_PATH'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.venv', 'Lib', 'site-packages',
                                            'PyQt5', 'Qt5', 'plugins')


class ReportGenerator(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Генератор отчетов')

        # Create main layout
        layout = QVBoxLayout()

        # QTextEdit for preview
        self.textEdit = QTextEdit(self)
        self.textEdit.setReadOnly(True)
        layout.addWidget(self.textEdit)

        # Buttons for generating and saving the report
        self.generateButton1 = QPushButton("Генерировать отчет 1", self)
        self.generateButton2 = QPushButton("Генерировать отчет 2", self)
        self.generateButton3 = QPushButton("Генерировать отчет 3", self)
        self.generateButton4 = QPushButton("Генерировать отчет 4", self)
        self.saveButton = QPushButton("Сохранить в PDF/ODF", self)

        layout.addWidget(self.generateButton1)
        layout.addWidget(self.generateButton2)
        layout.addWidget(self.generateButton3)
        layout.addWidget(self.generateButton4)
        layout.addWidget(self.saveButton)

        self.setLayout(layout)

        # Connect buttons to actions
        self.generateButton1.clicked.connect(lambda: self.generate_report(1))
        self.generateButton2.clicked.connect(lambda: self.generate_report(2))
        self.generateButton3.clicked.connect(lambda: self.generate_report(3))
        self.generateButton4.clicked.connect(lambda: self.generate_report(4))
        self.saveButton.clicked.connect(self.save_report)

    def generate_report(self, report_type):
        # Create a QTextDocument to hold the report content
        document = self.textEdit.document()
        cursor = QTextCursor(document)

        # Set font and styles for the document
        font = QFont("Calibri")

        # Title style (size 16, bold, center aligned)
        title_format = QTextCharFormat()
        title_format.setFont(QFont("Calibri", 16))
        title_format.setFontWeight(QFont.Bold)
        title_block_format = QTextBlockFormat()
        title_block_format.setAlignment(Qt.AlignCenter)

        # Ordinary text style (size 11, left aligned)
        ordinary_block_format = QTextBlockFormat()
        ordinary_block_format.setAlignment(Qt.AlignLeft)
        ordinary_format = QTextCharFormat()
        ordinary_format.setFont(QFont("Calibri", 11))

        # Bold text style (size 11, bold, center aligned)
        bold_block_format = QTextBlockFormat()
        bold_block_format.setAlignment(Qt.AlignCenter)
        bold_format = QTextCharFormat()
        bold_format.setFont(QFont("Calibri", 11))
        bold_format.setFontWeight(QFont.Bold)

        # picture
        picture_block_format = QTextBlockFormat()
        picture_block_format.setAlignment(Qt.AlignCenter)
        picture_format = QTextCharFormat()
        picture_format.setFont(QFont("Calibri", 11))

        # Clear the document before adding content
        document.clear()

        # Generate content for each report type
        if report_type == 1:
            cursor.setBlockFormat(title_block_format)
            cursor.setCharFormat(title_format)
            cursor.insertText("РЕФЕРАТ\n\n")

            cursor.setBlockFormat(ordinary_block_format)
            cursor.setCharFormat(ordinary_format)
            cursor.insertText("        Пояснительная записка содержит 109 страниц, 35 иллюстраций, 47 таблиц, 42 источника.\n\n")

            cursor.setBlockFormat(bold_block_format)
            cursor.setCharFormat(bold_format)
            cursor.insertText("БАЗА ДАННЫХ, ХРАНИЛИЩЕ ДАННЫХ, OLAP-КУБ, СИСТЕМА УПРАВЛЕНИЯ БАЗАМИ ДАННЫХ (СУБД), МНОГОПОТОЧНОЕ ПРОГРАММИРОВАНИЕ, ИНТЕРФЕЙС ПРИКЛАДНОГО ПРОГРАММИРОВАНИЯ.\n\n")

            cursor.setBlockFormat(ordinary_block_format)
            cursor.setCharFormat(ordinary_format)
            cursor.insertText(
                """        Объектом исследования является организация хранения данных вуза, относящихся к его учебной деятельности.
        Результатом проектирования является хранилище данных вуза.
        Проект выполнен с применением методики системного структурного анализа. Выполнен анализ методов хранения данных, проведен анализ предметной области – организации учебной деятельности вуза, определены потребности вуза в этом аспекте. Спроектировано и реализовано хранилище данных на основе комбинирования SQL базы данных и файлового хранилища, разработан сетевой API для доступа к данным. Проведено внедрение и опытная эксплуатация хранилища данных на базе СамГТУ.
        Были использованы технологии, языки программирования и протоколы: C++, Framework Qt, MS Access, HTTP, XML, JSON, многопоточное программирование.\n\n"""
            )

        elif report_type == 2:
            cursor.setBlockFormat(ordinary_block_format)
            cursor.setCharFormat(ordinary_format)
            cursor.insertText(
                """        PostgreSQL отличается от других РСУБД тем, что обладает объектно-ориентированным функционалом, в том числе полной поддержкой концепта ACID (Atomicity, Consistency, Isolation, Durability).
        Postgres отлично справляется с одновременной обработкой нескольких заданий. Поддержка конкурентности реализована с использованием MVCC (Multiversion Concurrency Control), что также обеспечивает совместимость с ACID [11, 12].\n\n"""
            )

            cursor.setBlockFormat(picture_block_format)
            self.insert_image(cursor, "resources/postgre.png")
            cursor.insertBlock()

            cursor.setBlockFormat(picture_block_format)
            cursor.setCharFormat(picture_format)
            cursor.insertText("Рисунок 2.5 – Взаимодействие прикладного процесса с PostreSQL\n\n")

            cursor.setBlockFormat(ordinary_block_format)
            cursor.setCharFormat(ordinary_format)
            cursor.insertText(
                """        В простых операциях чтения PostgreSQL уступает своим аналогам (MySQL). Также PostreSQL довольна сложна и не очень распространена на бесплатных хостингах. Однако, если приоритет стоит на надёжность и целостность данных или база данных должна выполнять сложные процедуры, PostgreSQL – хороший  выбор [13].
        Слабой стороной PostgreSQL – это быстрые операции чтения.
        
  Таблица 2.1 – Таблица сравнения реляционных БД""")

            # Create a table with 2 rows and 4 columns
            table = cursor.insertTable(2, 4)

            # Set the text for the first row (headers)
            cursor.insertText("Продукт")
            cursor.movePosition(QTextCursor.NextCell)
            cursor.insertText("Производитель")
            cursor.movePosition(QTextCursor.NextCell)
            cursor.insertText("Тип системы БД")
            cursor.movePosition(QTextCursor.NextCell)
            cursor.insertText("Поддержка объектов БД")

            # Move to the second row and set text
            cursor.movePosition(QTextCursor.NextRow)
            cursor.insertText("ORACLE SQL Server")
            cursor.movePosition(QTextCursor.NextCell)
            cursor.insertText("Oracle")
            cursor.movePosition(QTextCursor.NextCell)
            cursor.insertText("реляционная")
            cursor.movePosition(QTextCursor.NextCell)
            cursor.insertText("Таблицы, представления, индексы, хранимые процедуры, пользовательские функции, роли")
            cursor.movePosition(QTextCursor.Down)
            cursor.insertText("\n\n")

            cursor.setBlockFormat(ordinary_block_format)
            cursor.setCharFormat(ordinary_format)
            cursor.insertText(
                """Можно выделить четыре различных типов NoSQL баз данных [18].
                1.	Документоориентированне базы данных или хранилищам документов. Этот NoSQL тип баз даны предназначен для хранения, извлечения, обработки и управления документно ориентированной информацией. В этих базах данных обучно присутствует ключ, который связан со сложной структурой данных называемой документом.
                2.	Хранилища ключ-значение представляет собой коллекцию ключей, каждый из которых связан только с одним значением. Иногда этот тип баз данных называют словарем. Наверное, хранилища ключ значение является самым простым и распространенным типом, обеспечивающим максимальное быстродействие среди всех баз данных NoSQL.
                3.	Хранилища с широкими столбцами использует таблицы, строки и столбцы как в традиционных реляционных базах данных. Однако, имена и формат столбцов могут варьироваться в различных строках в пределах одной таблицы, что в корне отличает их отличие от реляционных сбаз данных.
                4.	Хранилища графов – графовая база данных использует графовые структуры для семантических запросов с узлами, ребрами и свойствами для представления и хранения данных.""")

        elif report_type == 3:
            cursor.setBlockFormat(ordinary_block_format)
            cursor.setCharFormat(ordinary_format)
            cursor.insertText("        По работе «Динамические формы» сделать генератор PDF анкет (1 файл – одна анкета) и сборщик всех анкет в один сводный отчет диаграммами различного типа (круговые и столбчатые), иллюстрирующие предпочтения пользователей, распределение по возрастам, …. И таблицами данных по которым подводились итоги, списками анкет … То есть сводный отчет. Обрабатываются все JSON файлы из выбранной папки. Отчет должен иметь презентабельный вид.")

        elif report_type == 4:
            cursor.setBlockFormat(ordinary_block_format)
            cursor.setCharFormat(ordinary_format)
            cursor.insertText("        По работе «Журнал успеваемости» сделать генератор PDF отчета группы по успеваемости: таблица успеваемости, графики распределений, например по дисциплинам и оценкам. Сводные таблицы со средними показателями. Выбор исходного файла через стандартное меню выбора файла. Отчет должен иметь презентабельный вид.")

    def save_report(self):
        # Open a file dialog for saving
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getSaveFileName(self, "Сохранить как", "", "PDF Files (*.pdf);;ODF Files (*.odt)")

        if file_path:
            # Generate the document for saving
            document = self.textEdit.document()

            # Create a printer object
            printer = QPrinter()
            if file_path.endswith(".pdf"):
                printer.setOutputFormat(QPrinter.PdfFormat)
            elif file_path.endswith(".odt"):
                printer.setOutputFormat(QPrinter.OdfFormat)

            printer.setOutputFileName(file_path)
            document.print_(printer)
            print("Документ сохранен: ", file_path)

    def insert_image(self, cursor, image_path):
        # Load the image using QImage
        image = QImage(image_path)

        if image.isNull():
            print("Image loading failed.")
            return

        # Create a QTextImageFormat to insert the image
        image_format = QTextImageFormat()
        image_format.setName(image_path)  # Set image path (or the image name)

        # Insert the image at the cursor position
        cursor.insertImage(image_format)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ReportGenerator()
    window.show()
    sys.exit(app.exec_())
