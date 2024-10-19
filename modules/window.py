import os, zipfile, shutil
from psd_tools import PSDImage
from threading import *

# QT Imports
from PySide6.QtGui import QColor, QBrush, QPixmap, QIcon, QBrush, QCursor
from PySide6.QtCore import QPoint, Signal, Qt, QRectF
from PySide6.QtWidgets import QFrame, QDialog, QMessageBox, QGridLayout, QGraphicsView, QGraphicsScene, QHBoxLayout, QGraphicsPixmapItem, QVBoxLayout, QLabel, QGroupBox, QFileDialog, QPushButton, QTableWidget, QTableWidgetItem

SCALE_FACTOR = 1.25

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

def zip_directory(directory_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), directory_path))

class PhotoViewer(QGraphicsView):
    coordinatesChanged = Signal(QPoint)

    def __init__(self, parent):
        super().__init__(parent)
        self._zoom = 0
        self._pinned = False
        self._empty = True
        self._scene = QGraphicsScene(self)
        self._photo = QGraphicsPixmapItem()
        self._photo.setShapeMode(QGraphicsPixmapItem.ShapeMode.BoundingRectShape)
        self._scene.addItem(self._photo)
        self.setScene(self._scene)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setBackgroundBrush(QBrush(QColor(209, 209, 209)))
        self.setFrameShape(QFrame.Shape.NoFrame)

    def hasPhoto(self):
        return not self._empty

    def resetView(self, scale=1):
        rect = QRectF(self._photo.pixmap().rect())
        if not rect.isNull():
            self.setSceneRect(rect)
            if (scale := max(1, scale)) == 1:
                self._zoom = 0
            if self.hasPhoto():
                unity = self.transform().mapRect(QRectF(0, 0, 1, 1))
                self.scale(1 / unity.width(), 1 / unity.height())
                viewrect = self.viewport().rect()
                scenerect = self.transform().mapRect(rect)
                factor = min(viewrect.width() / scenerect.width(),
                             viewrect.height() / scenerect.height()) * scale
                self.scale(factor, factor)
                if not self.zoomPinned():
                    self.centerOn(self._photo)
                self.updateCoordinates()

    def setPhoto(self, pixmap=None):
        if pixmap and not pixmap.isNull():
            self._empty = False
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            self._photo.setPixmap(pixmap)
        else:
            self._empty = True
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
            self._photo.setPixmap(QPixmap())
        if not (self.zoomPinned() and self.hasPhoto()):
            self._zoom = 0
        self.resetView(SCALE_FACTOR ** self._zoom)

    def zoomLevel(self):
        return self._zoom

    def zoomPinned(self):
        return self._pinned

    def setZoomPinned(self, enable):
        self._pinned = bool(enable)

    def zoom(self, step):
        zoom = max(0, self._zoom + (step := int(step)))
        if zoom != self._zoom:
            self._zoom = zoom
            if self._zoom > 0:
                if step > 0:
                    factor = SCALE_FACTOR ** step
                else:
                    factor = 1 / SCALE_FACTOR ** abs(step)
                self.scale(factor, factor)
            else:
                self.resetView()

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        self.zoom(delta and delta // abs(delta))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.resetView()

    def toggleDragMode(self):
        if self.dragMode() == QGraphicsView.DragMode.ScrollHandDrag:
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
        elif not self._photo.pixmap().isNull():
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)

    def updateCoordinates(self, pos=None):
        if self._photo.isUnderMouse():
            if pos is None:
                pos = self.mapFromGlobal(QCursor.pos())
            point = self.mapToScene(pos).toPoint()
        else:
            point = QPoint()
        self.coordinatesChanged.emit(point)

    def mouseMoveEvent(self, event):
        self.updateCoordinates(event.position().toPoint())
        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        self.coordinatesChanged.emit(QPoint())
        super().leaveEvent(event)

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
        self.table_box = self.compare_table()
        
        right_layout = QVBoxLayout()
        self.preview_pages = self.page_preview()
        self.export_box = self.export()
        right_layout.addWidget(self.preview_pages)
        right_layout.addWidget(self.export_box)

        footer = self.footer()

        self.table_box.setEnabled(False)
        self.export_box.setEnabled(False)
        self.preview_pages.setEnabled(False)

        # Main Layout
        main_layout = QGridLayout(self)
        main_layout.addLayout(header, 0, 0, 1, 2)
        main_layout.addWidget(self.table_box, 1, 0)
        main_layout.addLayout(right_layout, 1, 1)
        main_layout.addLayout(footer, 2, 0, 1, 2)

        main_layout.setColumnStretch(0, 1)
        main_layout.setColumnStretch(1, 3)

    def footer(self) -> QHBoxLayout:
        result = QHBoxLayout()
        label = QLabel("<i>by Mikis (Powered Developer) &lt;<a href='https://github.com/PoweredDeveloper'>https://github.com/PoweredDeveloper</a>&gt;</i>")
        label.setOpenExternalLinks(True)
        result.addWidget(QLabel('v0.2'))
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

        frame = QFrame()
        frame.setFixedSize(400, 600)

        viewer = PhotoViewer(frame)
        viewer.setPhoto(pixmap)
        viewer.setFixedSize(400, 600)

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
        export_btn.clicked.connect(self.export_file)
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

    def export_file(self) -> None:
        if self.export_path == '': return

        result_path = os.path.join(os.getcwd(), 'buffer', 'result')
        os.mkdir(result_path)

        for file in os.listdir(os.path.join(os.getcwd(), 'buffer', 'initial')):
            shutil.copy(os.path.join(os.getcwd(), 'buffer', 'initial', file), result_path)

        for file in os.listdir(os.path.join(os.getcwd(), 'buffer', 'corrected')):
            shutil.copy(os.path.join(os.getcwd(), 'buffer', 'corrected', file), result_path)

        zip_directory(result_path, 'result.zip')
        
        if os.path.exists(os.path.join(self.export_path, 'result.zip')):
            shutil.move(os.path.join(os.getcwd(), 'result.zip'), os.path.join(self.export_path, 'result.zip'))
        else:
            shutil.move(os.path.join(os.getcwd(), 'result.zip'), self.export_path)

        msg = QMessageBox()
        msg.setWindowTitle("Export")
        msg.setText("Archive was successfully exported")
        msg.setWindowIcon(QIcon(os.path.join(os.getcwd(), 'assets', 'icon.png')))
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()
        
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
        
        self.table_box.setEnabled(True)
        self.export_box.setEnabled(True)
        self.preview_pages.setEnabled(True)

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