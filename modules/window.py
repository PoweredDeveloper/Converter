import os, zipfile, shutil
from psd_tools import PSDImage
from threading import *

# QT Imports
from PySide6.QtGui import QColor, QBrush, QPixmap, QIcon
from PySide6.QtWidgets import QFrame, QDialog, QGridLayout, QHBoxLayout, QVBoxLayout, QLabel, QGroupBox, QFileDialog, QPushButton, QTableWidget, QTableWidgetItem

def cutstr(string: str, length: int = 30) -> str:
    return '...' + string[-length:] if len(string) > length else string

def get_filename(path: str) -> str:
    return path.split('/')[-1]

def read_zip(path: str, extension: str = '') -> list[str]:
    files = []
    with zipfile.ZipFile(path, 'r') as zip:
        for file in zip.filelist:
            if (file.filename.split('.')[-1] == extension) or extension == '':
                files.append(file.filename)

    return files

def extract_all(zip_path: str, end_directory: str, file_extension: str = '') -> None:
    with zipfile.ZipFile(zip_path, 'r') as zip:
        for file in zip.filelist:
            if (file.filename.split('.')[-1] == file_extension) or file_extension == '':
                zip.extract(file, end_directory)

def convert_psd(path: str, filename: str, converted_name: str) -> None:
    filepath = os.path.join(path, filename)
    if not os.path.exists(filepath): return

    psd = PSDImage.open(filepath)
    new_filepath = os.path.join(path, str(converted_name) + '.png')
    psd.composite().save(new_filepath)

def sort_pages(pages: list[str]) -> list[str]:
    pages.sort(key = lambda file: int(''.join(filter(str.isdigit, file))))
    return pages

def delete_with_extension(path: str, extension: str) -> None:
    for file in os.listdir(path):
        if file.split('.')[-1] == extension:
            os.remove(os.path.join(path, file))

class Window(QDialog):
    def __init__(self):
        super().__init__()
        # Class variables
        self.pages = { 'initial': [], 'corrected': [] }
        self.archives = { 'initial': '', 'corrected': '' }
        self.export_path = ''
        self.page = 1

        # Window Settings
        self.setWindowTitle('Manga PSD to PNG converter')
        self.setWindowIcon(QIcon(os.path.join(os.getcwd(), 'assets', 'icon.png')))

        # Layouts & Widgets
        header = self.header()
        table = self.compare_table()
        
        right_layout = QVBoxLayout()
        preview = self.page_preview()
        export = self.export()
        right_layout.addWidget(preview)
        right_layout.addWidget(export)

        footer = self.footer()

        # Main Layout
        main_layout = QGridLayout(self)
        main_layout.addLayout(header, 0, 0, 1, 2)
        main_layout.addWidget(table, 1, 0)
        main_layout.addLayout(right_layout, 1, 1)
        main_layout.addLayout(footer, 2, 0, 1, 2)

        main_layout.setColumnStretch(0, 1)
        main_layout.setColumnStretch(1, 3)

    def footer(self) -> QHBoxLayout:
        result = QHBoxLayout()
        label = QLabel("<i>by Mikis (Powered Developer) &lt;<a href='https://github.com/PoweredDeveloper'>https://github.com/PoweredDeveloper</a>&gt;</i>")
        label.setOpenExternalLinks(True)
        result.addWidget(QLabel('v0.1'))
        result.addStretch(1)
        result.addWidget(label)
        return result

    def header(self) -> QHBoxLayout:
        initial_select_btn = QPushButton("Select")
        initial_select_btn.clicked.connect(lambda: self.select_archive('initial', initial_select_btn))
        corrected_select_btn = QPushButton("Select")
        corrected_select_btn.clicked.connect(lambda: self.select_archive('corrected', corrected_select_btn))
        proceed_btn = QPushButton("Proceed   >")
        proceed_btn.clicked.connect(self.proceed_project)
        
        result = QHBoxLayout()
        result.addWidget(QLabel('Select Initial:'))
        result.addWidget(initial_select_btn)
        result.addWidget(QLabel('Select Corrected:'))
        result.addWidget(corrected_select_btn)
        result.addStretch(1)
        result.addWidget(proceed_btn)

        return result

    def select_archive(self, archive_type: str, button: QPushButton) -> None:
        dialog = QFileDialog(self)
        dialog.setWindowTitle("Choose ZIP-file containing pages")
        dialog.setDirectory(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') )
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setNameFilter("ZIP files (*.zip)")

        if dialog.exec():
            filenames = dialog.selectedFiles()
            if filenames:
                self.archives[archive_type] = filenames[0]
                button.setText(cutstr(get_filename(filenames[0])))
        
    def compare_table(self) -> QGroupBox:
        reload_btn = QPushButton("Reload")
        reload_btn.clicked.connect(self.update)

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setRowCount(10)
        self.table.horizontalHeader().setSectionResizeMode(self.table.horizontalHeader().ResizeMode.Fixed)
        self.table.setHorizontalHeaderLabels(['Initial', 'Corrected'])

        result = QGroupBox('Compare Archives')
        layout = QVBoxLayout(result)
        layout.addWidget(reload_btn)
        layout.addWidget(self.table)

        return result
    
    def update_table(self) -> None:
        if list(self.archives.values()).count('') > 0: return

        self.table.setRowCount(len(self.pages['initial']))
        for i in range(len(self.pages['initial'])):
            initial_item = QTableWidgetItem(self.pages['initial'][i])
            corrected_item = QTableWidgetItem(self.pages['corrected'][i])

            if str(self.pages['corrected'][i]) == 'N/A':
                corrected_item.setForeground(QBrush(QColor(244, 19, 72)))

            self.table.setItem(i, 0, initial_item)
            self.table.setItem(i, 1, corrected_item)

    def page_preview(self) -> QGroupBox:
        result = QGroupBox('Preview')

        # Buttons
        next_btn = QPushButton('>')
        next_btn.clicked.connect(lambda: self.update_pages(self.page + 1))
        previous_btn = QPushButton('<')
        previous_btn.clicked.connect(lambda: self.update_pages(self.page - 1))

        self.page_label = QLabel(".. / ..")

        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(previous_btn)
        buttons_layout.addStretch(1)
        buttons_layout.addWidget(self.page_label)
        buttons_layout.addStretch(1)
        buttons_layout.addWidget(next_btn)

        # Pages
        self.preview_initial = self.preview('blank')
        self.preview_corrected = self.preview('blank')

        self.pages_layout = QHBoxLayout()
        self.pages_layout.addWidget(self.preview_initial)
        self.pages_layout.addWidget(self.preview_corrected)

        # Layout
        vertical_layout = QVBoxLayout(result)
        vertical_layout.addLayout(buttons_layout)
        vertical_layout.addLayout(self.pages_layout)

        return result
    
    def preview(self, type: str, page: int = 1, color: QColor = QColor(209, 209, 209)) -> QFrame:
        filename = None
        if type != 'blank':
            for index in range(len(self.pages[type])):
                if type == 'initial' and index + 1 == page:
                    filename = self.pages[type][index]
                    break
                elif type == 'corrected' and index + 1 == page:
                    filename = self.pages[type][index] if self.pages[type][index] != 'N/A' else None
                    break

        if (filename == None and type != 'initial') or type == 'blank':
            pixmap = pixmap = QPixmap(400, 600)
            pixmap.fill(color)
        else:
            pixmap = QPixmap(os.path.join(os.getcwd(), 'buffer', type, filename))

        pixmap = pixmap.scaledToHeight(600)

        frame = QFrame()
        frame.setFixedSize(400, 600)

        page = QLabel(frame)
        page.setPixmap(pixmap)

        return frame
    
    def update_pages(self, page: int = 1) -> None:
        if page > len(self.pages['initial']) or page <= 0: return
        self.page = page

        self.page_label.setText(str(self.page) + ' / ' + str(len(self.pages['initial'])))

        if list(self.archives.values()).count('') <= 0:
            preview_initial = self.preview('initial', page)
            preview_corrected = self.preview('corrected', page)

            self.pages_layout.replaceWidget(self.preview_initial, preview_initial)
            self.pages_layout.replaceWidget(self.preview_corrected, preview_corrected)

            self.preview_initial = preview_initial
            self.preview_corrected = preview_corrected

    def export(self) -> QHBoxLayout:
        result = QGroupBox('Export')
        layout = QHBoxLayout(result)

        export_btn = QPushButton('Export')
        export_btn.setMinimumHeight(40)
        export_btn.setMinimumWidth(140)

        select_path_btn = QPushButton('Select')
        select_path_btn.clicked.connect(lambda: self.select_save_path(select_path_btn))

        layout.addWidget(QLabel('Save path: '))
        layout.addWidget(select_path_btn)
        layout.addStretch(1)
        layout.addWidget(export_btn)

        return result
    
    def select_save_path(self, button: QPushButton) -> None:
        dialog = QFileDialog(self)
        dialog.setWindowTitle("Choose Location, where Zip will be saved")
        dialog.setDirectory(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') )
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        dialog.setNameFilter("ZIP files (*.zip)")

        if dialog.exec():
            filenames = dialog.selectedFiles()
            if filenames:
                self.export_path = filenames[0]
                button.setText(cutstr(self.export_path, 90))
        
    def proceed_project(self) -> None:
        if list(self.archives.values()).count('') > 0: return

        self.pages = { 'initial': [], 'corrected': [] }
        self.page = 1

        buffer_dir = os.path.join(os.getcwd(), 'buffer')
        initial_dir = os.path.join(buffer_dir, 'initial')
        corrected_dir = os.path.join(buffer_dir, 'corrected')

        if os.path.exists(buffer_dir):
            shutil.rmtree(buffer_dir, ignore_errors = True)

        os.mkdir(buffer_dir)
        os.mkdir(initial_dir)
        os.mkdir(corrected_dir)

        extract_all(self.archives['initial'], initial_dir, 'psd')
        extract_all(self.archives['corrected'], corrected_dir, 'psd')

        self.load_files()

        self.pages['initial'] = sort_pages(os.listdir(os.path.join(os.getcwd(), 'buffer', 'initial')))
        self.pages['corrected'] = sort_pages(os.listdir(os.path.join(os.getcwd(), 'buffer', 'corrected')))
        self.pages['corrected'] = [self.pages['initial'][i] if self.pages['corrected'].count(self.pages['initial'][i]) > 0 else 'N/A' for i in range(len(self.pages['initial']))]
        
        self.update()

    def update(self) -> None:
        self.update_table()
        self.update_pages()

    def load_files(self) -> None:
        buffer_dir = os.path.join(os.getcwd(), 'buffer')
        initial_dir = os.path.join(buffer_dir, 'initial')
        corrected_dir = os.path.join(buffer_dir, 'corrected')

        # NEED TO OPTIMIZE
        convert_initial = [Thread(target = convert_psd, args = (initial_dir, file, int(''.join(filter(str.isdigit, file))))) for file in os.listdir(initial_dir)]
        convert_corrected = [Thread(target = convert_psd, args = (corrected_dir, file, int(''.join(filter(str.isdigit, file))))) for file in os.listdir(corrected_dir)]
        
        for thread in convert_initial:
            thread.start()

        for thread in convert_initial:
            thread.join()

        for thread in convert_corrected:
            thread.start()

        for thread in convert_corrected:
            thread.join()

        for file in os.listdir(initial_dir):
            if file.split('.')[-1] == 'psd':
                os.remove(os.path.join(initial_dir, file))

        for file in os.listdir(corrected_dir):
            if file.split('.')[-1] == 'psd':
                os.remove(os.path.join(corrected_dir, file))