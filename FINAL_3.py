import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QLabel,
)


def input_1(df1):
    df1 = df1[['축약코드', '거래수량']]
    df1.rename(columns={'축약코드': '종목코드'}, inplace=True)
    # df1['거래수량'] = df1['거래수량'].apply(lambda x: x.replace(',', ''))
    df1 = df1.astype({'거래수량': 'float'})
    df1 = df1.groupby('종목코드').sum()
    return df1


def input_2(df2):
    df2 = df2[['축약코드', '거래수량']]
    df2.rename(columns={'축약코드': '종목코드'}, inplace=True)
    # df2['거래수량'] = df2['거래수량'].apply(lambda x: x.replace(',', ''))
    df2 = df2.astype({'거래수량': 'float'})
    df2 = df2.groupby('종목코드').sum()
    return df2


def input_3(df3):
    df3 = df3[['종목코드', '상환수량']]
    # df3['상환수량'] = df3['상환수량'].apply(lambda x: x.replace(',', ''))
    df3 = df3.astype({'상환수량': 'float'})
    df3 = df3.groupby('종목코드').sum()
    return df3


def input_4(df4):
    df4 = df4[['종목코드', '대차수량']]
    # df4['대차수량'] = df4['대차수량'].apply(lambda x: x.replace(',', ''))
    df4 = df4.astype({'대차수량': 'float'})
    df4 = df4.groupby('종목코드').sum()
    return df4


def input_final(df1, df2, df3, df4):

    a = input_1(df1)
    b = input_2(df2)
    c = input_3(df3)
    d = input_4(df4)

    temp_1 = pd.merge(a, b, how='outer', on='종목코드')
    temp_2 = pd.merge(c, d, how='outer', on='종목코드')
    result = pd.merge(temp_1, temp_2, how='outer', on='종목코드')

    result.replace(np.NaN, 0, inplace=True)

    result['차이'] = (result['거래수량_x'] - result['거래수량_y']) - (result['상환수량'] - result['대차수량'])
    result['차이절대값'] = abs(result['차이'])

    result = result.sort_values('차이', ascending=True)

    result['종목코드'] = result.index

    result = result[['종목코드'] + [col for col in result.columns if col != '종목코드']]

    return result


class ExcelSheetLoaderAndMerger(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

        self.excel_file = None  # Store the selected Excel file path
        self.df1 = pd.DataFrame()
        self.df2 = pd.DataFrame()
        self.df3 = pd.DataFrame()
        self.df4 = pd.DataFrame()
        self.merged_df = pd.DataFrame()

    def initUI(self):
        self.setWindowTitle('Excel Sheet Loader and Merger')
        self.setGeometry(100, 100, 400, 400)

        layout = QVBoxLayout()

        self.select_button = QPushButton('Select Excel File', self)
        self.select_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.select_button)

        self.load_button = QPushButton('Load Sheets and Merge', self)
        self.load_button.clicked.connect(self.load_sheets_and_merge)
        layout.addWidget(self.load_button)

        self.export_button = QPushButton('Export Merged DataFrame', self)
        self.export_button.clicked.connect(self.export_merged_dataframe)
        layout.addWidget(self.export_button)

        self.result_label = QLabel('', self)
        layout.addWidget(self.result_label)

        self.central_widget = QWidget()
        self.central_widget.setLayout(layout)
        self.setCentralWidget(self.central_widget)

    def open_file_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        excel_file, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)", options=options
        )

        if excel_file:
            self.result_label.setText(f'Selected Excel File: {excel_file}')
            self.excel_file = excel_file

    def load_sheets_and_merge(self):
        if not self.excel_file:
            self.result_label.setText('Please select an Excel file first.')
            return

        try:
            excel_data = pd.ExcelFile(self.excel_file)
            self.df1 = pd.read_excel(excel_data, sheet_name='62051_차입상환')
            self.df2 = pd.read_excel(excel_data, sheet_name='62051_차입')
            self.df3 = pd.read_excel(excel_data, sheet_name='13014_상환')
            self.df4 = pd.read_excel(excel_data, sheet_name='13014_대여')

            # Merge the DataFrames into one merged DataFrame
            self.merged_df = input_final(self.df1, self.df2, self.df3, self.df4)
            self.result_label.setText('Loaded all sheets and merged into one DataFrame.')

        except Exception as e:
            self.result_label.setText(f'Error loading sheets: {str(e)}')

    def export_merged_dataframe(self):
        if self.merged_df.empty:
            self.result_label.setText('No merged DataFrame to export.')
            return

        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        export_file, _ = QFileDialog.getSaveFileName(
            self, "Export Merged DataFrame", "", "Excel Files (*.xlsx *.xls)", options=options
        )

        if export_file:
            try:
                self.merged_df.to_excel(export_file, index=False)
                self.result_label.setText(f'Merged DataFrame exported to: {export_file}')
            except Exception as e:
                self.result_label.setText(f'Error exporting DataFrame: {str(e)}')


def main():
    app = QApplication(sys.argv)
    window = ExcelSheetLoaderAndMerger()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
