from layout import Ui_MainWindow
from PyQt5 import QtWidgets

if __name__ == '__main__':
    class Window(QtWidgets.QMainWindow):
        def __init__(self):
            super().__init__()
            self.components = Ui_MainWindow()
            self.components.setupUi(self)
            self.show()
            self.components.start_button.clicked.connect(lambda: print(self.components.path_input.text()))
            # TODO file selector
            self.components.use_id_button.toggled.connect(lambda: print('status: ' + str(self.components.use_id_button.isChecked())))
            # TODO editable elements
            # TODO group disable

            def temp():
                dl = QtWidgets.QFileDialog()
                dl.setAcceptMode(QtWidgets.QFileDialog.AcceptSave)
                dl.setNameFilter('Excel File (*.xlsx)')

                if dl.exec_():
                    filenames = dl.selectedFiles()
                    print(filenames)

            self.components.open_button.clicked.connect(temp)
            self.disable_group = [self.components.sheet_name_input,
                                  self.components.pub_title_input,
                                  self.components.lib_title_input,
                                  self.components.letter_only_button,
                                  self.components.no_space_button,
                                  self.components.strict_button,
                                  self.components.correct_color_box,
                                  self.components.ambiguous_color_box,
                                  self.components.wrong_color_box]

            self.id_group = [self.components.pub_id_input,
                             self.components.lib_id_input]

            self.color_group = [self.components.correct_color_box,
                                self.components.ambiguous_color_box,
                                self.components.wrong_color_box]

            COLOR_ITEM = ['aliceblue', 'antiquewhite', 'aqua', 'aquamarine', 'azure', 'beige', 'bisque', 'black',
                          'blanchedalmond', 'blue', 'blueviolet', 'brown', 'burlywood', 'cadetblue', 'chartreuse',
                          'chocolate', 'coral', 'cornflowerblue', 'cornsilk', 'crimson', 'cyan', 'darkblue',
                          'darkcyan', 'darkgoldenrod', 'darkgray', 'darkgreen', 'darkgrey', 'darkkhaki', 'darkmagenta',
                          'darkolivegreen', 'darkorange', 'darkorchid', 'darkred', 'darksalmon', 'darkseagreen',
                          'darkslateblue', 'darkslategray', 'darkslategrey', 'darkturquoise', 'darkviolet', 'deeppink',
                          'deepskyblue', 'dimgray', 'dimgrey', 'dodgerblue', 'firebrick', 'floralwhite', 'forestgreen',
                          'fuchsia', 'gainsboro', 'ghostwhite', 'gold', 'goldenrod', 'gray', 'green', 'greenyellow',
                          'grey', 'honeydew', 'hotpink', 'indianred', 'indigo', 'ivory', 'khaki', 'lavender',
                          'lavenderblush', 'lawngreen', 'lemonchiffon', 'lightblue', 'lightcoral', 'lightcyan',
                          'lightgoldenrodyellow', 'lightgray', 'lightgreen', 'lightgrey', 'lightpink', 'lightsalmon',
                          'lightseagreen', 'lightskyblue', 'lightslategray', 'lightslategrey', 'lightsteelblue',
                          'lightyellow', 'lime', 'limegreen', 'linen', 'magenta', 'maroon', 'mediumaquamarine',
                          'mediumblue', 'mediumorchid', 'mediumpurple', 'mediumseagreen', 'mediumslateblue',
                          'mediumspringgreen', 'mediumturquoise', 'mediumvioletred', 'midnightblue', 'mintcream',
                          'mistyrose', 'moccasin', 'navajowhite', 'navy', 'oldlace', 'olive', 'olivedrab', 'orange',
                          'orangered', 'orchid', 'palegoldenrod', 'palegreen', 'paleturquoise', 'palevioletred',
                          'papayawhip', 'peachpuff', 'peru', 'pink', 'plum', 'powderblue', 'purple', 'rebeccapurple',
                          'red', 'rosybrown', 'royalblue', 'saddlebrown', 'salmon', 'sandybrown', 'seagreen',
                          'seashell', 'sienna', 'silver', 'skyblue', 'slateblue', 'slategray', 'slategrey', 'snow',
                          'springgreen', 'steelblue', 'tan', 'teal', 'thistle', 'tomato', 'turquoise', 'violet',
                          'wheat', 'white', 'whitesmoke', 'yellow', 'yellowgreen']

            for color_box in self.color_group:
                color_box.addItems(COLOR_ITEM)

            self.components.correct_color_box.setCurrentIndex(self.components.correct_color_box.findText('green'))
            self.components.ambiguous_color_box.setCurrentIndex(self.components.ambiguous_color_box.findText('yellow'))
            self.components.wrong_color_box.setCurrentIndex(self.components.wrong_color_box.findText('red'))

            def change_disable_status():
                dis_flag = self.components.use_user_enter_button.isChecked()
                id_flag = self.components.use_id_button.isChecked()
                for element in self.disable_group:
                    element.setDisabled(not dis_flag)

                for element in self.id_group:
                    element.setDisabled(not (id_flag and dis_flag))

            def change_id_status():
                dis_flag = self.components.use_user_enter_button.isChecked()
                id_flag = self.components.use_id_button.isChecked()

                for element in self.id_group:
                    element.setDisabled(not (id_flag and dis_flag))

            change_disable_status()

            self.components.use_user_enter_button.clicked.connect(change_disable_status)
            self.components.use_id_button.clicked.connect(change_id_status)

            # TODO if user use self defined data, give a warninig


            def start_helper():
                file_path = self.components.path_input.text()



    import sys
    app = QtWidgets.QApplication(sys.argv)
    main_win = Window()
    app.exec_()