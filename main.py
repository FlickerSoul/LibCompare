#
#
#   Created By @Flicker_Soul
#   Date: Jan. 26 2020
#
#   Main File For LibCompareApp
#
#

from PyQt5 import QtWidgets
from layout import Ui_MainWindow
import data
import logging
from logging import handlers
import traceback

import sys
from pathlib import Path
import time


FORMAT_STRING = '%(asctime)s -- %(levelname)s: %(message)s'
FORMATTER = logging.Formatter(FORMAT_STRING)
logging.basicConfig(format=FORMAT_STRING, level=logging.DEBUG)
LOGGER = logging.getLogger('Logger')

LOG_PATH = str(Path.home()) + '/Documents/LibCompare/'
Path(LOG_PATH).mkdir(parents=True, exist_ok=True)
LOG_PATH += 'log'

with open(LOG_PATH, 'w') as f:
    f.close()

FILE_LOGGER_HANDLER = handlers.TimedRotatingFileHandler(
    filename=LOG_PATH, backupCount=5
)
FILE_LOGGER_HANDLER.setFormatter(FORMATTER)
FILE_LOGGER_HANDLER.setLevel(logging.INFO)
#
STREAM_LOGGER_HANDLER = logging.StreamHandler()
STREAM_LOGGER_HANDLER.setLevel(logging.INFO)
STREAM_LOGGER_HANDLER.setFormatter(FORMATTER)

# LOGGER.addHandler(FILE_LOGGER_HANDLER)
# LOGGER.addHandler(STREAM_LOGGER_HANDLER)


def show_critical_dialog(txt, info):
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Critical)
    msg.setText(txt)
    msg.setInformativeText(info)
    msg.setWindowTitle('Error')
    msg.exec_()


def show_warning_dialog(txt, info):
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setText(txt)
    msg.setInformativeText(info)
    msg.setWindowTitle('Warning')
    msg.exec_()


def show_info_dialog(txt, info):
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Information)
    msg.setText(txt)
    msg.setInformativeText(info)
    msg.setWindowTitle('Info')
    msg.exec_()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        LOGGER.info('========== New Start ==========')
        super().__init__()
        self.components = Ui_MainWindow()
        self.components.setupUi(self)

        self.components.path_input.setText('/Users/flicker_soul/Documents/Python Demo/LibWork/test/test.xlsx')
        self.components.pub_end_row.setText('1405')
        self.components.lib_end_row.setText('1573')

        LOGGER.debug('Initialize Components')

        self.show()

        LOGGER.debug('Show Window')

        def call_file_chooser():
            fd = QtWidgets.QFileDialog()
            fd.setFileMode(QtWidgets.QFileDialog.AnyFile)
            fd.setNameFilter('Excel File (*.xlsx)')

            if fd.exec_():
                filenames = fd.selectedFiles()
                if len(filenames) > 1:
                    LOGGER.warning('Chose Many Files Only The First One Will Be Used')
                elif len(filenames) < 1:
                    LOGGER.info('No File Chose')
                    return

                path = filenames[0]
                self.components.path_input.setText(path)
                LOGGER.debug('Set File Path To ' + path)
                return

            LOGGER.info('No File Chose')

        def call_file_saver():
            fd = QtWidgets.QFileDialog()
            fd.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)

            if fd.exec_():
                filenames = fd.selectedFiles()
                if len(filenames) > 1:
                    LOGGER.warning('Multiple Save Target Chose, Only The First One Will Be Used')
                elif len(filenames) < 1:
                    LOGGER.info('No File Chose')
                    return ''

                path = filenames[0]
                return path

            LOGGER.info('No File Chose')
            return ''

        self.components.open_button.clicked.connect(call_file_chooser)

        LOGGER.debug('Setup File Chooser')

        self.disable_group = [self.components.sheet_name_input,
                              self.components.pub_title_input,
                              self.components.lib_title_input,
                              # self.components.letter_only_button,
                              # self.components.no_space_button,
                              # self.components.strict_button,
                              self.components.correct_color_box,
                              self.components.ambiguous_color_box,
                              self.components.wrong_color_box]

        self.components.letter_only_button.setDisabled(True)
        self.components.no_space_button.setDisabled(True)
        self.components.strict_button.setDisabled(True)

        LOGGER.debug('Defined Disable Group')

        self.id_group = [self.components.pub_id_input,
                         self.components.lib_id_input]

        LOGGER.debug('Defined Id Group')

        self.color_group = [self.components.correct_color_box,
                            self.components.ambiguous_color_box,
                            self.components.wrong_color_box]

        LOGGER.debug('Defined Color Group')

        COLOR_SET = data.output_color_name()

        for color_box in self.color_group:
            color_box.addItems(COLOR_SET)

        self.components.correct_color_box.setCurrentIndex(self.components.correct_color_box.findText('green'))
        self.components.ambiguous_color_box.setCurrentIndex(self.components.ambiguous_color_box.findText('yellow'))
        self.components.wrong_color_box.setCurrentIndex(self.components.wrong_color_box.findText('red'))

        LOGGER.debug('Setup Color Choices')

        def change_disable_status():
            dis_flag = self.components.use_user_enter_button.isChecked()
            id_flag = self.components.use_id_button.isChecked()
            for element in self.disable_group:
                element.setDisabled(not dis_flag)

            for element in self.id_group:
                element.setDisabled(not (id_flag and dis_flag))

        change_disable_status()
        self.components.use_user_enter_button.clicked.connect(change_disable_status)

        LOGGER.debug('Setup Use User Data Button')

        def change_id_status():
            dis_flag = self.components.use_user_enter_button.isChecked()
            id_flag = self.components.use_id_button.isChecked()

            for element in self.id_group:
                element.setDisabled(not (id_flag and dis_flag))

        self.components.use_id_button.clicked.connect(change_id_status)

        LOGGER.debug('Setup Use Id Button')

        def start_helper():

            is_using_id_compare = self.components.use_id_button.isChecked()
            is_using_user_defined_data = self.components.use_user_enter_button.isChecked()

            if not is_using_id_compare:
                show_warning_dialog(
                    'You Are Not Using ID Compare',
                    'Using Title Only Will Compromise Accuracy'
                )

            if is_using_user_defined_data:
                show_warning_dialog(
                    'You Are Using Self-Defined Data',
                    'Some Unknown Error May Occur'
                )

            file_path = self.components.path_input.text()
            work_book_wrapper = None
            try:
                work_book_wrapper = data.WorkBookWrapper(work_book_file_path=file_path)
            except Exception:
                LOGGER.debug(msg='Invalid File', stack_info=traceback.format_exc())
                show_critical_dialog('Invalid File', 'Must Use Valid Excel File (*.xlsx)')
                return False, None

            pub_id_col = data.WorkBookWrapper.PUB_ID_COL
            lib_id_col = data.WorkBookWrapper.LIB_ID_COL

            if is_using_id_compare:
                pub_id_col = self.components.pub_id_input.text()
                lib_id_col = self.components.lib_id_input.text()

            LOGGER.debug(('' if is_using_id_compare else 'Not ') + 'Use Id Compare')
            LOGGER.debug('\n' + 'Publisher\'s Id Col is ' + pub_id_col + '\n' + 'Lib\'s Id Col is ' + lib_id_col + '\n')

            data_sheet_name = data.WorkBookWrapper.DEFAULT_DATA_SHEET_NAME

            pub_title_col = data.WorkBookWrapper.PUB_TITLE_COL
            lib_title_col = data.WorkBookWrapper.LIB_TITLE_COL

            is_using_no_weird_letters = False
            is_using_no_space = False
            is_using_strict_mode = False

            correct_color = data.WorkBookWrapper.DEFAULT_CORRECT_COLOR
            ambiguous_color = data.WorkBookWrapper.DEFAULT_AMBIGUOUS_COLOR
            wrong_color = data.WorkBookWrapper.DEFAULT_WRONG_COLOR

            if is_using_user_defined_data:
                data_sheet_name = self.components.sheet_name_input.text()
                pub_title_col = self.components.pub_title_input.text()
                lib_title_col = self.components.lib_title_input.text()
                is_using_no_weird_letters = self.components.letter_only_button.isChecked()
                is_using_no_space = self.components.no_space_button.isChecked()
                is_using_strict_mode = self.components.strict_button.isChecked()
                correct_color = self.components.correct_color_box.currentText()
                ambiguous_color = self.components.ambiguous_color_box.currentText()
                wrong_color = self.components.wrong_color_box.currentText()

            LOGGER.debug(('' if is_using_user_defined_data else 'Not ') + 'Using User Defined Data')
            LOGGER.debug('\n' +  'Data Sheet Name is ' + data_sheet_name + '\n' +
                         'Publisher\'s Title Col is ' + pub_title_col + '\n' +
                         'Lib\'s Title Col is ' + lib_title_col + '\n' +
                         'Using Weird letters? ' + str(is_using_no_weird_letters) + '\n' +
                         'Using No Space Mode? ' + str(is_using_no_space) + '\n' +
                         'Using Strict Mode? ' + str(is_using_strict_mode) + '\n' +
                         'Correct Color: ' + correct_color + '\n' +
                         'Ambiguous Color: ' + ambiguous_color + '\n' +
                         'Wrong Color: ' + wrong_color + '\b')

            pub_end_row = self.components.pub_end_row.text()
            lib_end_row = self.components.lib_end_row.text()

            LOGGER.debug('\n' + 'Publisher\'s End Row is ' + pub_end_row + '\n' +
                         'Lib\'s End Row is ' + lib_end_row + '\n')

            LOGGER.debug('Wrapping Up Data')

            try:
                work_book_wrapper.set_pub_id_col(self.components.pub_id_input.text())
                work_book_wrapper.set_lib_id_col(self.components.lib_id_input.text())
                work_book_wrapper.set_data_sheet(data_sheet_name)
                work_book_wrapper.set_pub_title_col(pub_title_col)
                work_book_wrapper.set_lib_title_col(lib_title_col)
                work_book_wrapper.set_no_weird_letter_flag(is_using_no_weird_letters)
                work_book_wrapper.set_no_space_flag(is_using_no_space)
                work_book_wrapper.set_strict_mode_flag(is_using_strict_mode)
                work_book_wrapper.set_correct_color(correct_color)
                work_book_wrapper.set_ambiguous_color(ambiguous_color)
                work_book_wrapper.set_wrong_color(wrong_color)
                work_book_wrapper.set_pub_end_row(pub_end_row)
                work_book_wrapper.set_lib_end_row(lib_end_row)
            except AssertionError:
                LOGGER.critical(msg='Not Valid Letters',
                                stack_info=traceback.format_exc())
                show_critical_dialog(
                    'Not Valid Letters',
                    'You Must Use English Letters'
                )
            except KeyError:
                LOGGER.critical(msg='No Data Sheet As ' + data_sheet_name + 'Exists',
                                stack_info=traceback.format_exc())
                show_critical_dialog(
                    'No Data Sheet As ' + data_sheet_name + 'Exists',
                    'Your Data Sheet Input Is Not Valid'
                )
            except ValueError:
                LOGGER.critical(msg='Invalid End Row Input',
                                stack_info=traceback.format_exc())
                show_critical_dialog(
                    'Invalid End Row Input',
                    'You Need To Input A Valid Number As End Row'
                )
            else:
                LOGGER.info('Start Working')
                return work_book_wrapper.work(), work_book_wrapper

            return False, None

        def start():
            self.components.start_button.setDisabled(True)
            LOGGER.debug('Disable Start Button')

            show_info_dialog('Choose Output File Afterward',
                             'Please Select Output File Path Afterward')

            save_path = call_file_saver()
            if not save_path:
                LOGGER.warning('Nothing Will Be Saved')
                show_warning_dialog(
                    'No Save File Selected',
                    'Nothing Will Be Saved'
                )
            LOGGER.info('The Save Path Is ' + save_path)
            LOGGER.info('Save Work Book as ' + save_path)

            result = False
            wb = None

            start_time = time.time()
            result, wb = start_helper()
            LOGGER.info('Running Time: ' + str(time.time() - start_time) + 'S')

            try:
                pass
            except Exception:
                LOGGER.critical(msg='Unknown Error', stack_info=traceback.format_exc())
                show_critical_dialog(
                    'Unknown Error',
                    'Unknown Error Occurred; Please Contact Developer'
                )
            finally:
                self.components.start_button.setDisabled(False)
                LOGGER.debug('Unlock Start Button')

            LOGGER.info('SUCCESSFULLY EXECUTED: ' + str(result))
            if result:
                show_info_dialog(
                    'Comparing Completed',
                    'SUCCESSFULLY EXECUTED: ' + str(result).upper()
                )

                if wb is not None and save_path:
                    wb.work_book.save(save_path)
                    LOGGER.info('File Saved As ' + save_path)
            else:
                show_critical_dialog(
                    'Comparing Not Complete',
                    'Something Went Wrong'
                )

        self.components.start_button.clicked.connect(start)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main_win = MainWindow()
    app.exec_()
