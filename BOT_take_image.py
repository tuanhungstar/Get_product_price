# ... (all imports and code at the beginning of BOT_take_image.py are unchanged) ...
import sys
import json
import os
import clipboard
import io
import base64
import PIL.Image

from PyQt6.QtGui import QPixmap, QColor, QFont, QPainter, QPen, QIcon, QPolygonF, QCursor, QAction
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QVariant, QObject, QSize, QPoint, QRegularExpression,QRect,QDateTime, QTimer, QPointF, QTime
from PyQt6 import QtWidgets, QtGui, QtCore

from PyQt6 import QtCore, QtGui, QtWidgets # Import everything needed from PyQt6
from PyQt6.QtGui import QIcon # <-- ADD THIS LINE
from PyQt6.QtWidgets import QMessageBox # Explicitly import QMessageBox for clarity
# from PyQt6.QtCore import pyqtSignal # Already imported from PyQt6.QtCore implicitly with the above line
from PIL.ImageQt import ImageQt
from pymsgbox import confirm
import pygetwindow as gw
import pyautogui
# --- TKINTER IMPORTS REMOVED ---
import re # Needed for RPA step parsing if you re-add it, or for file list parsing
import time # Added this import, as it was used but not imported in Old_utility
from PyQt6.QtCore import pyqtSlot

# ... (The Old_utility class is unchanged) ...
class Old_utility:
    def __int__(self):
        
        pass
    def base64_pgn(self,text):
        image_data = base64.b64decode(text)
        img_file = PIL.Image.open(io.BytesIO(image_data))    
        return img_file
    def check_item_existing(self,current_file):
    
        with open('Click_image/{}.txt'.format(current_file)) as json_file:
            image_file = json.load(json_file)
        for key, data in image_file.items():
            image_file = self.base64_pgn(data)
            location=None
            try:
                location= pyautogui.locateCenterOnScreen(image_file,grayscale=True, confidence=0.98)
            except:
                pass
            if location!=None:
                return True
        return False
    def left_click(self, current_file, offset_x=0, offset_y=0, confidence=0.92):
        location = None
        txt_path = current_file if current_file.endswith('.txt') else f'Click_image/{current_file}.txt'
        if not os.path.exists(txt_path) and os.path.exists(f'{current_file}.txt'):
            txt_path = f'{current_file}.txt'
            
        if os.path.exists(txt_path):
            try:
                with open(txt_path, 'r', encoding='utf-8') as json_file:
                    image_data_dict = json.load(json_file)
                for key, data in image_data_dict.items():
                    img_obj = self.base64_pgn(data)
                    try:
                        location = pyautogui.locateCenterOnScreen(img_obj, grayscale=True, confidence=confidence)
                        if location is not None:
                            pyautogui.click(location.x + offset_x, location.y + offset_y)
                            return 'left_click_done'
                    except Exception:
                        pass
            except Exception:
                pass

        # Direct image file fallback (.png, .jpg, etc.)
        img_candidates = [
            current_file,
            f'Click_image/{current_file}',
            f'Click_image/{current_file}.png',
            f'Click_image/{current_file}.jpg'
        ]
        for img_p in img_candidates:
            if os.path.exists(img_p):
                try:
                    location = pyautogui.locateCenterOnScreen(img_p, grayscale=True, confidence=confidence)
                    if location is not None:
                        pyautogui.click(location.x + offset_x, location.y + offset_y)
                        return 'left_click_done'
                except Exception:
                    pass

        return f'left_click_done fail: {current_file}'
    def activate_window(self,title):
        loop=0
        full_tile =''
        while title not in full_tile:
            try:
                need_avtive = gw.getWindowsWithTitle(title)[0]
                #need_avtive.maximize()
                need_avtive.activate()
                full_tile = gw.getActiveWindow().title
                if title in full_tile:
                    #print(full_tile)
                    return 'done'
            except:
                #print ('try times:' , str(loop))
                loop +=1
                if loop >1000:
                    return 'fail'
    def close_window(self,title):
        loop=0
        full_tile =''
        while title not in full_tile:
            try:
                need_avtive = gw.getWindowsWithTitle(title)[0]
                need_avtive.close()
                all_window = gw.getAllTitles()
                wd_closed =None
                for each_window in all_window:
                    if title in full_tile:
                        wd_closed =False
                        break
                    else:
                        wd_closed =True
                if wd_closed==True:
                    return 'closed window'
            except:
                #print ('try times:' , str(loop))
                loop +=1
                if loop >1000 and wd_closed!=True:
                    return "Can't close {} window".format(title)   
    
    def waiting_window_close(self,window_name,timeout=5):
        '''
        Bot check if window is not exit
        return window_closed or return 'window_not_close'
        '''
        start_time = time.time()
        
        while int(time.time()-start_time) < timeout:
            all_window = gw.getAllTitles()
            result =  'window_closed'
            for each_window in all_window:
                if window_name in each_window:
                    result =  'window_not_close'
                    break # exit for look, get list of windows again
            if result !=  'window_not_close':
                return 'window_closed'  
        return 'window_not_close'
    
    def check_win_title_exits(self,title,timeout=5):
        win = None
        st = time.time()
        time_run=0
        while win is None and time_run<timeout:
            all_window = gw.getAllTitles()
            for each_window in all_window:
                if title in each_window:
                    win = each_window
            if win !=None:
                return True
            ed = time.time()
            time_run = ed-st
        return False    
    def check_image_exits(self,image_check,timeout=5):
        win = False
        st = time.time()
        time_run=0
        while win is False and time_run<timeout:
            win = self.check_item_existing(image_check)
            if win !=False:
                return True
            ed = time.time()
            time_run = ed-st
        return False        
        
    def wait_image_disappear(self,image_check,timeout=5):
        
        win = True
        st = time.time()
        time_run=0
        while win is True and time_run<timeout:
            win = self.check_item_existing(image_check)
            if win ==False:
                return True
            ed = time.time()
            time_run = ed-st
        return False 

# ... (The Ui_MainWindow_Base class is unchanged) ...
class Ui_MainWindow_Base(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(679, 300)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(parent=self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 671, 291))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        
        # --- MODIFICATION: Increased X for middle column ---
        MIDDLE_COL_X = 140 # Was 130
        
        self.label_folder = QtWidgets.QLabel(parent=self.tab)
        self.label_folder.setGeometry(QtCore.QRect(10, 10, 111, 16)) # Left col
        self.label_folder.setObjectName("label_folder")
        
        self.conf_list_cob_folder = QtWidgets.QComboBox(parent=self.tab)
        self.conf_list_cob_folder.setGeometry(QtCore.QRect(MIDDLE_COL_X, 10, 231, 22)) # Middle col
        self.conf_list_cob_folder.setObjectName("conf_list_cob_folder")
        self.conf_list_cob_folder.addItem("")
        self.conf_list_cob_folder.setItemText(0, "")

        self.label = QtWidgets.QLabel(parent=self.tab)
        self.label.setGeometry(QtCore.QRect(10, 40, 111, 16)) # Left col
        self.label.setObjectName("label")
        
        self.conf_list_cob_img = QtWidgets.QComboBox(parent=self.tab)
        self.conf_list_cob_img.setGeometry(QtCore.QRect(MIDDLE_COL_X, 40, 231, 22)) # Middle col
        self.conf_list_cob_img.setObjectName("conf_list_cob_img")
        self.conf_list_cob_img.addItem("")
        self.conf_list_cob_img.setItemText(0, "")
        
        self.KPI_botcomment_txt = QtWidgets.QListWidget(parent=self.tab)
        self.KPI_botcomment_txt.setGeometry(QtCore.QRect(MIDDLE_COL_X, 70, 231, 91)) # Middle col
        self.KPI_botcomment_txt.setObjectName("KPI_botcomment_txt")
        
        # --- BUTTONS (Left Column) ---
        self.Start_BOT_butt = QtWidgets.QPushButton(parent=self.tab)
        self.Start_BOT_butt.setGeometry(QtCore.QRect(10, 70, 120, 28))
        self.Start_BOT_butt.setObjectName("Start_BOT_butt")
        
        self.Save_img_butt = QtWidgets.QPushButton(parent=self.tab)
        self.Save_img_butt.setGeometry(QtCore.QRect(10, 100, 120, 28))
        self.Save_img_butt.setObjectName("Save_img_butt")
        
        self.delete_img_butt = QtWidgets.QPushButton(parent=self.tab)
        self.delete_img_butt.setGeometry(QtCore.QRect(10, 130, 120, 28))
        self.delete_img_butt.setObjectName("delete_img_butt")

        self.create_folder_butt = QtWidgets.QPushButton(parent=self.tab)
        self.create_folder_butt.setGeometry(QtCore.QRect(10, 160, 120, 28))
        self.create_folder_butt.setObjectName("create_folder_butt")
        
        self.exit_BOT_butt = QtWidgets.QPushButton(parent=self.tab)
        self.exit_BOT_butt.setGeometry(QtCore.QRect(10, 190, 120, 28))
        self.exit_BOT_butt.setObjectName("exit_BOT_butt")
        # --- END BUTTONS ---

        # --- Middle Column Continued ---
        self.label_3 = QtWidgets.QLabel(parent=self.tab)
        self.label_3.setGeometry(QtCore.QRect(MIDDLE_COL_X, 170, 111, 16)) # Middle col
        self.label_3.setObjectName("label_3")
        
        self.file_name_txt = QtWidgets.QTextEdit(parent=self.tab)
        self.file_name_txt.setGeometry(QtCore.QRect(MIDDLE_COL_X, 190, 231, 31)) # Middle col
        self.file_name_txt.setToolTip("")
        self.file_name_txt.setObjectName("file_name_txt")
        # --- END MIDDLE COLUMN ---

        # --- RIGHT SIDE WIDGETS (MOVED RIGHT previously) ---
        RIGHT_COL_X = 390 # Keep the previous adjustment for gap
        
        self.click_img = QtWidgets.QLabel(parent=self.tab)
        self.click_img.setGeometry(QtCore.QRect(RIGHT_COL_X, 40, 281, 91))
        self.click_img.setAutoFillBackground(True)
        self.click_img.setFrameShape(QtWidgets.QFrame.Shape.Box)
        self.click_img.setText("")
        self.click_img.setObjectName("click_img")
        
        self.label_2 = QtWidgets.QLabel(parent=self.tab)
        self.label_2.setGeometry(QtCore.QRect(RIGHT_COL_X, 10, 101, 16))
        self.label_2.setObjectName("label_2")
        
        self.Test_img_butt = QtWidgets.QPushButton(parent=self.tab)
        self.Test_img_butt.setGeometry(QtCore.QRect(RIGHT_COL_X, 140, 120, 28))
        self.Test_img_butt.setObjectName("Test_img_butt")
        
        self.label_test_result = QtWidgets.QLabel(parent=self.tab)
        self.label_test_result.setGeometry(QtCore.QRect(RIGHT_COL_X + 130, 140, 151, 31)) # X = 520
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        self.label_test_result.setFont(font)
        self.label_test_result.setText("")
        self.label_test_result.setObjectName("label_test_result")
        
        self.label_active_window = QtWidgets.QLabel(parent=self.tab)
        self.label_active_window.setGeometry(QtCore.QRect(RIGHT_COL_X, 170, 151, 16))
        self.label_active_window.setObjectName("label_active_window")
        
        self.conf_list_cob_window = QtWidgets.QComboBox(parent=self.tab)
        self.conf_list_cob_window.setGeometry(QtCore.QRect(RIGHT_COL_X, 190, 281, 31))
        self.conf_list_cob_window.setObjectName("conf_list_cob_window")
        self.conf_list_cob_window.addItem("")
        self.conf_list_cob_window.setItemText(0, "")
        # --- END RIGHT SIDE WIDGETS ---
        
        self.tabWidget.addTab(self.tab, "")

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Hphung BOT"))
        self.label_folder.setText(_translate("MainWindow", "Folder"))
        self.label.setText(_translate("MainWindow", "Image File"))
        self.Start_BOT_butt.setText(_translate("MainWindow", "New Screenshot"))
        self.Save_img_butt.setText(_translate("MainWindow", "Save_image"))
        self.delete_img_butt.setText(_translate("MainWindow", "Delete image"))
        self.create_folder_butt.setText(_translate("MainWindow", "New Folder"))
        self.exit_BOT_butt.setText(_translate("MainWindow", "Exit"))
        self.label_2.setText(_translate("MainWindow", "Image"))
        self.Test_img_butt.setText(_translate("MainWindow", "Test Image"))
        self.label_3.setText(_translate("MainWindow", "Image File Name"))
        self.label_active_window.setText(_translate("MainWindow", "Select active Window"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Screenshot"))

# ... (The DummySignal class is unchanged) ...
class DummySignal:
    def emit(self, *args, **kwargs):
        pass

# ... (The QtScreenshotter class is unchanged) ...
class QtScreenshotter(QtWidgets.QWidget):
    """
    A native PyQt6 fullscreen widget for capturing a region of the screen.
    Replaces the tkinter-based Screenshoot_gui.
    """
    # Signal that emits the screenshot as a base64 string when done
    screenshotTaken = QtCore.pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.start_pos = None
        self.selection_rect = QtCore.QRect()

        # Configure the widget to be a borderless, semi-transparent overlay
        self.setWindowFlags(
            QtCore.Qt.WindowType.FramelessWindowHint |
            QtCore.Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet("background-color: transparent;")
        self.setCursor(QtCore.Qt.CursorShape.CrossCursor)
        self.showFullScreen()

    def paintEvent(self, event):
        """Draws the semi-transparent overlay and the selection rectangle."""
        painter = QtGui.QPainter(self)
        
        # Draw the semi-transparent black overlay
        painter.fillRect(self.rect(), QtGui.QColor(0, 0, 0, 70))
        
        if self.start_pos:
            # "Punch a hole" in the overlay by clearing the selection area
            painter.setCompositionMode(QtGui.QPainter.CompositionMode.CompositionMode_Clear)
            painter.fillRect(self.selection_rect, QtCore.Qt.GlobalColor.transparent)
            
            # Reset composition mode to draw the border
            painter.setCompositionMode(QtGui.QPainter.CompositionMode.CompositionMode_SourceOver)
            painter.setPen(QtGui.QPen(QtCore.Qt.GlobalColor.red, 2, QtCore.Qt.PenStyle.SolidLine))
            painter.drawRect(self.selection_rect)
            
        painter.end()

    def mousePressEvent(self, event):
        """Stores the starting position of the drag."""
        self.start_pos = event.position().toPoint()
        self.selection_rect = QtCore.QRect(self.start_pos, self.start_pos)
        event.accept()

    def mouseMoveEvent(self, event):
        """Updates the selection rectangle as the mouse is dragged."""
        if self.start_pos:
            current_pos = event.position().toPoint()
            self.selection_rect = QtCore.QRect(self.start_pos, current_pos).normalized()
            self.update()  # Trigger a repaint
        event.accept()

    def mouseReleaseEvent(self, event):
        """Finalizes the selection and triggers the screenshot."""
        if self.start_pos:
            self.end_pos = event.position().toPoint()
            self.selection_rect = QtCore.QRect(self.start_pos, self.end_pos).normalized()
            
            # Hide the overlay *before* taking the screenshot
            self.hide()
            
            # Use a short timer to ensure the window is hidden
            QtCore.QTimer.singleShot(100, self.take_screenshot)
        event.accept()

    def keyPressEvent(self, event):
        """Allows canceling the screenshot with the Escape key."""
        if event.key() == QtCore.Qt.Key.Key_Escape:
            self.screenshotTaken.emit("") # Emit empty string to signal cancellation
            self.close()

    def take_screenshot(self):
        """Captures the selected region, base64-encodes it, and emits the signal."""
        if not self.selection_rect.isValid() or self.selection_rect.width() < 1 or self.selection_rect.height() < 1:
            self.screenshotTaken.emit("")
            self.close()
            return

        # Map the local widget rectangle to global screen coordinates
        global_rect = QtCore.QRect(
            self.mapToGlobal(self.selection_rect.topLeft()),
            self.mapToGlobal(self.selection_rect.bottomRight())
        )

        left = global_rect.x()
        top = global_rect.y()
        width = global_rect.width()
        height = global_rect.height()

        image = pyautogui.screenshot(region=(int(left), int(top), int(width), int(height)))
        
        base64_image = ""
        clipboard.copy('')
        if image is not None:
            image = image.convert('L') # Convert to grayscale
            buffer = io.BytesIO()
            image.save(buffer, format='PNG')
            base64_image = base64.b64encode(buffer.getvalue()).decode('utf-8')
            clipboard.copy(base64_image)
            
        self.screenshotTaken.emit(base64_image)
        self.close() # Close the snipping widget

# ... (The base64_pgn function is unchanged) ...
def base64_pgn(text):
    image_data = base64.b64decode(text)
    img_file = PIL.Image.open(io.BytesIO(image_data))
    return img_file

# ... (The get_list_file function is unchanged) ...
def get_list_file(root_folder, key_word, extension):
    if not os.path.exists(root_folder):
        os.makedirs(root_folder)
        return []

    files = os.listdir(root_folder)
    file_all = []
    for file in files:
        if extension in file and key_word in file:
            file_all.append(os.path.join(root_folder, file))
    return file_all

# ... (The QThread_Start_BOT class is unchanged) ...
class QThread_Start_BOT(QtCore.QThread):
    command_signal = QtCore.pyqtSignal(str) # Use QtCore.pyqtSignal

    def __init__(self, parent=None):
        super().__init__(parent)

    def thread_input(self, input_data):
        self.get_arrg = input_data

    def run(self):
        received_command = self.get_arrg
        #print(received_command)
        all_step = received_command.splitlines()
        for steps in all_step:
            if not steps.strip():
                continue
            try:
                parts = steps.split(': ', 1)
                if len(parts) < 2:
                    print(f"Warning: Skipping malformed step: {steps}")
                    continue

                action_parts = parts[1].split('>>')
                if len(action_parts) < 2:
                    #print(f"Warning: Skipping malformed action in step: {steps}")
                    continue

                class_name = action_parts[0].strip()
                func_and_arg = action_parts[1].strip()

                match = re.match(r'(\w+)\((.*)\)', func_and_arg)
                if not match:
                    #print(f"Warning: Skipping malformed function call in step: {steps}")
                    continue

                sub_func = match.group(1)
                agr = match.group(2).strip().replace("'", "")

                # IMPORTANT: In this context, globals() might not contain your main app's classes.
                # This RPA execution is highly coupled to its own environment.
                # If these are meant to interact with your main app's modules,
                # you'd need a more sophisticated execution context passing.
                if class_name in globals() and hasattr(globals()[class_name], sub_func):
                    m = globals()[class_name]()
                    func = getattr(m, sub_func)
                    func(agr)
                else:
                    #print(f"Error: Class '{class_name}' or function '{sub_func}' not found for step: {steps}")
                    pass
            except Exception as e:
                print(f"Error processing RPA step '{steps}': {e}")


# --- Main Application Window Class ---
# ... (The MainWindow class __init__ and all other methods are unchanged) ...
class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow_Base):
    # NEW SIGNAL: To be emitted when a screenshot is saved
    screenshotSaved = QtCore.pyqtSignal(str)

    def __init__(self, statup_image_param: str):
        super().__init__()
        self.setupUi(self)
        # --- ADD THIS: Countdown Timer Setup ---
        self.countdown_timer = QTimer()
        self.countdown_timer.timeout.connect(self.on_countdown_tick)
        self.countdown_val = 5
        # --- ADD THESE LINES ---
        # Define the path relative to this script
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # Goes up one level from my_lib
        icon_path = os.path.join(base_dir, "app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        # --- END OF ADDED LINES ---
        
        self.move(1030, 0)
        
        self.current_pic = '' # Initialize instance attribute
        self.img_data = {}    # Initialize instance attribute
        self.snipper = None   # <-- ADD THIS: To hold the screenshot widget
        
        self.statup_image = statup_image_param # Store as instance attribute

        # --- FOLDER AND FILE LOGIC ---
        self.click_image_dir = 'Click_image'
        os.makedirs(self.click_image_dir, exist_ok=True)
        
        self._all_image_paths = [] # Holds relative paths e.g., "Folder/img1", "img_root"
        self._folder_list = []     # Holds just folder names e.g., "Folder"
        
        self.load_folders_and_files()
        self.populate_folder_dropdown()
        self.filter_image_list() # Initial population of file list
        # --- END FOLDER LOGIC ---

        # Check if startup image is already in the combo box
        if self.statup_image:
            # Set folder dropdown if path contains a folder
            if '/' in self.statup_image:
                folder_name = self.statup_image.split('/')[0]
                self.conf_list_cob_folder.setCurrentText(folder_name) # This triggers filter_image_list
            
            # Set the file in the (now filtered) file list
            self.conf_list_cob_img.setCurrentText(self.statup_image)

        self.conf_list_img() # Call this after setting the current text to load preview

        all_window_titles = gw.getAllTitles()
        for each_window_title in all_window_titles:
            if len(each_window_title) > 2:
                self.conf_list_cob_window.addItem(each_window_title)

        # --- Connect UI Elements ---
        self.Start_BOT_butt.clicked.connect(self.Start_BOT)
        self.exit_BOT_butt.clicked.connect(self.exit_BOT) # This will be the closing action
        self.Save_img_butt.clicked.connect(self.Save_img)
        self.delete_img_butt.clicked.connect(self.delete_img)
        self.create_folder_butt.clicked.connect(self.create_new_folder) # NEW
        self.Test_img_butt.clicked.connect(self._test_image_actual) 

        self.KPI_botcomment_txt.clicked.connect(self.KPI_txt)
        
        self.conf_list_cob_folder.currentIndexChanged.connect(self.filter_image_list) # NEW
        self.conf_list_cob_img.currentIndexChanged.connect(self.conf_list_img)

    # --- NEW FOLDER/FILE METHODS ---
    def load_folders_and_files(self):
        """Scans Click_image dir and populates internal lists of folders and relative image paths."""
        self._all_image_paths = []
        self._folder_list = []
        
        for root, dirs, files in os.walk(self.click_image_dir):
            # Add folders
            if root != self.click_image_dir:
                folder_name = os.path.relpath(root, self.click_image_dir).replace(os.sep, '/')
                self._folder_list.append(folder_name)
                
            # Add files
            for file in files:
                if file.lower().endswith(".txt"):
                    full_path = os.path.join(root, file)
                    relative_path = os.path.relpath(full_path, self.click_image_dir)
                    # Get path without extension and standardize separators
                    path_no_ext = os.path.splitext(relative_path)[0].replace(os.sep, '/')
                    self._all_image_paths.append(path_no_ext)
                    
        self._all_image_paths.sort()
        self._folder_list.sort()

    def populate_folder_dropdown(self):
        """Fills the folder dropdown with 'All Folders' and detected folder names."""
        self.conf_list_cob_folder.blockSignals(True)
        self.conf_list_cob_folder.clear()
        self.conf_list_cob_folder.addItem("All Folders")
        self.conf_list_cob_folder.addItems(self._folder_list)
        self.conf_list_cob_folder.blockSignals(False)

    def filter_image_list(self):
        """Filters the image file dropdown based on the selected folder."""
        selected_folder = self.conf_list_cob_folder.currentText()
        
        self.conf_list_cob_img.blockSignals(True)
        self.conf_list_cob_img.clear()
        self.conf_list_cob_img.addItem("---Add New---") # Add this as the first item
        filtered_list = []
        if selected_folder == "All Folders":
            filtered_list = self._all_image_paths
        else:
            prefix = selected_folder + '/'
            # Include files in the root folder (no prefix)
            root_files = [p for p in self._all_image_paths if '/' not in p]
            # Include files in the selected subfolder
            folder_files = [p for p in self._all_image_paths if p.startswith(prefix)]
            filtered_list = root_files + folder_files
            
        self.conf_list_cob_img.addItems(sorted(filtered_list))
        self.conf_list_cob_img.blockSignals(False)
        
        # Trigger a refresh of the image preview
        if self.conf_list_cob_img.count() > 0:
            self.conf_list_cob_img.setCurrentIndex(0)
            self.conf_list_img()
        else:
            self.conf_list_img() # This will clear the preview

    def create_new_folder(self):
        """Prompts user to create a new subfolder in Click_image."""
        folder_name, ok = QtWidgets.QInputDialog.getText(self, 'New Folder', 'Enter new folder name:')
        
        if ok and folder_name:
            # Sanitize folder name
            folder_name_safe = re.sub(r'[\\/:*?"<>|]', '', folder_name).strip()
            if not folder_name_safe or folder_name_safe == "All Folders":
                self.show_info_messagebox("Invalid folder name.")
                return

            new_folder_path = os.path.join(self.click_image_dir, folder_name_safe)
            
            if not os.path.exists(new_folder_path):
                try:
                    os.makedirs(new_folder_path)
                    self.show_info_messagebox(f"Folder '{folder_name_safe}' created.")
                    # Refresh folder list
                    self._folder_list.append(folder_name_safe)
                    self._folder_list.sort()
                    self.populate_folder_dropdown()
                    # Select the new folder
                    self.conf_list_cob_folder.setCurrentText(folder_name_safe)
                except Exception as e:
                    self.show_info_messagebox(f"Could not create folder: {e}")
            else:
                self.show_info_messagebox(f"Folder '{folder_name_safe}' already exists.")
                self.conf_list_cob_folder.setCurrentText(folder_name_safe)
    # --- END NEW METHODS ---

    # --- UI Slot Methods (MODIFIED) ---
    @pyqtSlot()
    def Start_BOT(self):
        # 1. Check File Logic (Keep your existing logic)
        current_selection = self.conf_list_cob_img.currentText()
        proceed = True
        
        if current_selection and current_selection != "---Add New---":
            reply = QMessageBox.question(self, 'Check file', 'Do you want to add to current file?',
                                         QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
            if reply != QMessageBox.StandardButton.Ok:
                proceed = False

        if proceed:
            # 2. Setup Countdown
            self.countdown_val = 5
            self.Start_BOT_butt.setText(str(self.countdown_val))
            self.Start_BOT_butt.setEnabled(False) # Disable button so user can't click twice
            self.countdown_timer.start(1000) # Fire every 1000ms (1 second)

    def on_countdown_tick(self):
        """Called every second by the timer."""
        self.countdown_val -= 1
        
        if self.countdown_val > 0:
            # Update button text
            self.Start_BOT_butt.setText(str(self.countdown_val))
        else:
            # Countdown finished
            self.countdown_timer.stop()
            self.Start_BOT_butt.setText("New Screenshot") # Reset text
            self.Start_BOT_butt.setEnabled(True) # Re-enable button
            
            # Trigger the actual screenshot process
            self.launch_screenshot_tool()

    def launch_screenshot_tool(self):
        """Minimizes window and starts the snipper."""
        self.setWindowState(QtCore.Qt.WindowState.WindowMinimized)
        
        # Use a singleShot timer to give the window time to minimize animation 
        # before the screenshot overlay appears (avoids ghosting)
        QTimer.singleShot(300, self._show_snipper)

    def _show_snipper(self):
        """Internal helper to actually show the snipping widget."""
        self.snipper = QtScreenshotter()
        self.snipper.screenshotTaken.connect(self.on_screenshot_finished)
        self.snipper.show()
    @pyqtSlot(str)
    def on_screenshot_finished(self, base64_image):
        """
        This slot is called when the QtScreenshotter widget emits its
        screenshotTaken signal.
        """
        self.setWindowState(QtCore.Qt.WindowState.WindowNoState) # Restore window
        self.raise_()
        self.activateWindow() # Bring to front
        
        self.current_pic = base64_image  # Update instance attribute

        if len(self.current_pic) > 20:
            pic_png = base64_pgn(self.current_pic)
            qimage = ImageQt(pic_png)
            #pixmap = QtGui.QPixmap.fromImage(qimage)
            pixmap = self.resize_qimage_and_create_qpixmap(qimage)
            self.click_img.setPixmap(pixmap)
        
        # Clean up the snipper instance
        if self.snipper:
            self.snipper.deleteLater()
            self.snipper = None
        Old_utility().activate_window('Take and Manage')
    @pyqtSlot()
    def Save_img(self):
        if len(self.current_pic) <= 50:
            self.show_info_messagebox('No image data to save. Take a screenshot first.')
            return

        selected_folder = self.conf_list_cob_folder.currentText()
        
        # --- REQUIREMENT 6: Check folder selection ---
        if selected_folder == "All Folders":
            # Check if there are any root files (no folder)
            root_files_exist = any('/' not in p for p in self._all_image_paths)
            if not root_files_exist and not self._folder_list:
                # Special case: No folders exist, allow saving to root
                selected_folder = "" # Save to root
            else:
                self.show_info_messagebox('Please select a specific folder (or "New Folder") before saving.')
                
                return
        # ---

        file_name = self.file_name_txt.toPlainText().strip()
        if not file_name:
            self.show_info_messagebox('Please Enter a File Name.')
            return

        # Sanitize file_name for filesystem safety
        file_name_safe = re.sub(r'[\\/:*?"<>|]', '', file_name)
        file_name_safe = file_name_safe.replace(' ', '_')
        if not file_name_safe:
            self.show_info_messagebox('Invalid file name. Please use alphanumeric characters.')
            self.__init__()
        file_name = file_name_safe 

        # --- New Path Logic ---
        if selected_folder: # Not root
            relative_path_no_ext = f"{selected_folder}/{file_name}"
        else: # Root
            relative_path_no_ext = file_name
            
        file_path_txt = os.path.join(self.click_image_dir, relative_path_no_ext + ".txt")
        # ---

        file_name_exists_in_all_paths = relative_path_no_ext in self._all_image_paths
        
        # Determine if we should load existing data
        load_existing = False
        if file_name_exists_in_all_paths:
            reply = QMessageBox.question(self, 'File Exists', f'File "{relative_path_no_ext}" already exists. Do you want to "Add" to it, or "Overwrite" it?',
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Cancel:
                return
            elif reply == QMessageBox.StandardButton.No: # Overwrite
                self.img_data = {} # Clear existing data
            elif reply == QMessageBox.StandardButton.Yes: # Add
                load_existing = True
        else: # New file name, no conflict
            self.img_data = {} # Start with empty data for new file

        if load_existing:
            try:
                with open(file_path_txt, 'r') as json_file:
                    self.img_data = json.load(json_file)
            except (FileNotFoundError, json.JSONDecodeError):
                self.img_data = {} # If file is missing/corrupt, start new

        max_idx = -1
        if self.img_data:
            for k in self.img_data.keys():
                parts = k.rsplit('_', 1)
                if len(parts) > 1:
                    try:
                        idx = int(parts[-1])
                        max_idx = max(max_idx, idx)
                    except ValueError:
                        pass
        new_key_idx = max_idx + 1

        new_image_key = f'{file_name}_{new_key_idx}'
        self.img_data[new_image_key] = self.current_pic

        try:
            # Ensure the subfolder exists (it should, but check)
            os.makedirs(os.path.dirname(file_path_txt), exist_ok=True)
            
            with open(file_path_txt, 'w') as outfile:
                json.dump(self.img_data, outfile, indent=4)
            self.show_info_messagebox(f"Image saved as '{new_image_key}' in '{relative_path_no_ext}.txt'")

            # If the file_name was new, add it to our lists
            if not file_name_exists_in_all_paths:
                self._all_image_paths.append(relative_path_no_ext)
                self._all_image_paths.sort()
                self.filter_image_list() # Refresh the file list
                
            self.conf_list_cob_img.setCurrentText(relative_path_no_ext)
            self.conf_list_img() # Refresh the list widget
            self.KPI_botcomment_txt.setCurrentRow(self.KPI_botcomment_txt.count() - 1)
            
            # Emit signal that a screenshot was saved
            self.screenshotSaved.emit(relative_path_no_ext) # Emit the relative path

        except Exception as e:
            self.show_info_messagebox(f"Error saving image: {e}")
    @pyqtSlot()
    def delete_img(self):
        selected_item = self.KPI_botcomment_txt.currentItem()
        if not selected_item:
            self.show_info_messagebox('No image selected to delete.')
            return

        item_key_to_delete = selected_item.text()
        file_name_with_path = self.conf_list_cob_img.currentText()
        if not file_name_with_path:
            self.show_info_messagebox('No image file selected.')
            return

        reply = QMessageBox.question(self, 'Delete Image', f'Do you want to delete "{item_key_to_delete}" from "{file_name_with_path}.txt"?',
                                     QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        
        if reply == QMessageBox.StandardButton.Ok:
            if item_key_to_delete in self.img_data:
                del self.img_data[item_key_to_delete]
                self.KPI_botcomment_txt.takeItem(self.KPI_botcomment_txt.row(selected_item))
                
                file_path_txt = os.path.join(self.click_image_dir, file_name_with_path + ".txt")

                try:
                    if self.img_data:
                        # Save the file with the item removed
                        with open(file_path_txt, 'w') as outfile:
                            json.dump(self.img_data, outfile, indent=4)
                        self.show_info_messagebox(f"Image '{item_key_to_delete}' deleted successfully.")
                        if self.KPI_botcomment_txt.count() > 0:
                            self.KPI_botcomment_txt.setCurrentRow(0)
                            self.KPI_txt()
                        else: 
                            self.current_pic = '' 
                            self.click_img.clear()
                            self.screenshotSaved.emit("")
                    else: 
                        # If img_data became empty, remove the file
                        os.remove(file_path_txt)
                        self.show_info_messagebox(f"Image '{item_key_to_delete}' deleted and file '{file_name_with_path}.txt' removed as it is now empty.")
                        self.current_pic = ''
                        self.click_img.clear()
                        
                        # Remove from internal list and refresh GUI
                        if file_name_with_path in self._all_image_paths:
                            self._all_image_paths.remove(file_name_with_path)
                        self.filter_image_list() # This will remove it from conf_list_cob_img
                        
                        # Check if folder is now empty and remove it
                        folder_name = os.path.dirname(file_path_txt)
                        if folder_name != self.click_image_dir and not os.listdir(folder_name):
                            os.rmdir(folder_name)
                            self.load_folders_and_files() # Full reload
                            self.populate_folder_dropdown()
                            self.filter_image_list()
                        
                        self.screenshotSaved.emit(self.conf_list_cob_img.currentText()) # Emit new selection
                            
                except OSError as e:
                    #print (f"BOT Take Image Module: Delete Image Error: {e}")
                    self.show_info_messagebox(f"Error removing file '{file_name_with_path}.txt': {e}")
                except Exception as e:
                    #print (f"BOT Take Image Module: Delete Image Error: {e}")
                    self.show_info_messagebox(f"Error saving updated image file: {e}")
            else:
                self.show_info_messagebox('Selected image not found in data.')
    @pyqtSlot()
    def conf_list_img(self):
        self.KPI_botcomment_txt.clear()
        current_file_name = self.conf_list_cob_img.currentText() # This is now a relative path
        self.label_test_result.setText('')
        
        # --- ADD THIS BLOCK to handle the new item ---
        if current_file_name == "---Add New---":
            self.click_img.clear()
            self.current_pic = ''
            self.img_data = {}
            self.file_name_txt.setText("") # Clear file name box to prompt for new
            return # Stop processing
        # --- END OF ADDED BLOCK ---
        
        if current_file_name:
            file_path_txt = os.path.join(self.click_image_dir, current_file_name + ".txt")
            try:
                with open(file_path_txt, 'r') as json_file:
                    self.img_data = json.load(json_file) # Update instance attribute
                if self.img_data:
                    for imge_name, data in self.img_data.items():
                        self.KPI_botcomment_txt.addItem(imge_name)
                    first_key = list(self.img_data.keys())[0] if self.img_data else None
                    if first_key:
                        pic_png = base64_pgn(self.img_data[first_key])
                        qimage = ImageQt(pic_png)
                        #pixmap = QtGui.QPixmap.fromImage(qimage)
                        pixmap = self.resize_qimage_and_create_qpixmap(qimage)
                        self.click_img.setPixmap(pixmap)
                        self.current_pic = self.img_data[first_key] # Update instance attribute
                        items = self.KPI_botcomment_txt.findItems(first_key, QtCore.Qt.MatchFlag.MatchExactly)
                        if items:
                            self.KPI_botcomment_txt.setCurrentItem(items[0])
                else:
                    self.click_img.clear()
                    self.current_pic = ''
            except FileNotFoundError:
                # This can happen if filter_image_list() clears the list
                self.img_data = {}
                self.click_img.clear()
                self.current_pic = ''
            except json.JSONDecodeError:
                self.show_info_messagebox(f"Error decoding JSON from '{current_file_name}.txt'. File might be corrupted.")
                self.img_data = {}
                self.click_img.clear()
                self.current_pic = ''
        else:
            self.KPI_botcomment_txt.clear()
            self.click_img.clear()
            self.current_pic = ''
        
        # Set file name box to just the base name, not the folder path
        base_file_name = os.path.basename(current_file_name)
        self.file_name_txt.setText(base_file_name)

    @pyqtSlot()
    def KPI_txt(self):
        selected_item = self.KPI_botcomment_txt.currentItem()
        self.label_test_result.setText('')

        if selected_item:
            item_key = selected_item.text()
            if item_key in self.img_data:
                self.current_pic = self.img_data[item_key]
                pic_png = base64_pgn(self.current_pic)
                qimage = ImageQt(pic_png)
                #pixmap = QtGui.QPixmap.fromImage(qimage)
                pixmap = self.resize_qimage_and_create_qpixmap(qimage)
                self.click_img.setPixmap(pixmap)
            else:
                self.show_info_messagebox(f"Image data for key '{item_key}' not found.")
                self.click_img.clear()
                self.current_pic = ''
        else:
            self.click_img.clear()
            self.current_pic = ''
            
    def resize_qimage_and_create_qpixmap(self,qimage_input, percentage=98):
        if qimage_input.isNull():
            return QPixmap()
        new_width = int(qimage_input.width() * (percentage / 100))
        new_height = int(qimage_input.height() * (percentage / 100))
        return QPixmap.fromImage(qimage_input.scaled(QSize(new_width, new_height), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        
    @pyqtSlot()
    def _test_image_actual(self): # Renamed the actual test logic
        if len(self.current_pic) > 50:
            status = self.text_located(self.current_pic)
            if status is not None:
                self.label_test_result.setText('Image Found')
            else:
                self.label_test_result.setText('Image NOT Found')
        else:
            self.label_test_result.setText('No image data to test.')

    def text_located(self, im_base64_string):
        img_file = base64_pgn(im_base64_string)
        location = None
        try:
            location = pyautogui.locateCenterOnScreen(img_file, confidence=0.9)
        except Exception as e:
            #print(f"Error during image location: {e}")
            pass
        if location is not None:
            pyautogui.moveTo(location.x, location.y)
        return location
    @pyqtSlot()
    def exit_BOT(self):
        # Emit the currently selected relative path
        current_file_path = self.conf_list_cob_img.currentText()
        self.screenshotSaved.emit(current_file_path if current_file_path else "")
        self.close()

    def show_bot_comment(self, val):
        self.KPI_botcomment_txt.addItem(str(val))
        self.KPI_botcomment_txt.scrollToBottom()

    # -----------------------------------------------------------------
    # --- THIS IS THE MODIFIED FUNCTION ---
    # -----------------------------------------------------------------
    def show_info_messagebox(self, text, title="Information"):
        # --- FIX: Pass 'self' as the parent to the QMessageBox ---
        # This ensures the message box is "owned" by the screenshot window
        # and prevents it from closing the entire application.
        msg = QMessageBox(parent=self)
        # --- END FIX ---
        
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setText(text)
        msg.setWindowTitle(title)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        return msg.exec()
    # -----------------------------------------------------------------
    # --- END OF MODIFICATION ---
    # -----------------------------------------------------------------

if __name__ == "__main__":
    # Example usage for standalone testing
    app = QtWidgets.QApplication(sys.argv)
    # When running standalone, you need a dummy startup image name
    ex = MainWindow('default_screenshot') 
    ex.show() # <-- ADD THIS to actually show the window
    sys.exit(app.exec())
