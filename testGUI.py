from PyQt6.QtWidgets import (
    QMainWindow,
    QTextEdit,
    QFileDialog,
    QApplication,
    QPushButton,
    QHBoxLayout,
    QWidget,
    QVBoxLayout,
    QLineEdit
)
from PyQt6.QtGui import QIcon, QAction
from pathlib import Path
import sys

import pandas as pd
import numpy as np
import collections


class GUI(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setGeometry(300, 300, 350, 100)
        self.setWindowTitle('Automated Fundamental Analysis')

        self.generalLayout = QVBoxLayout()
        centralWidget = QWidget(self)
        centralWidget.setLayout(self.generalLayout)
        self.setCentralWidget(centralWidget)

        openFile = QAction(QIcon('open.png'), 'Open', self)
        openFile.setShortcut('Ctrl+O')
        openFile.setStatusTip('Open new File')
        openFile.triggered.connect(self.showDialog)

        downloadButton = QPushButton("Download CSV")
        downloadButton.clicked.connect(self.downloadFile)

        self.display = QLineEdit()
        self.display.setReadOnly(True)
        # self.display.setText("Downloaded Successfully!")

        layout = QHBoxLayout()
        layout.addWidget(downloadButton)
        layout.addWidget(self.display)
        self.generalLayout.addLayout(layout)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&Upload Report')
        fileMenu.addAction(openFile)


    def showDialog(self):
        home_dir = str(Path.home())
        fname = QFileDialog.getOpenFileName(self, 'Open file', home_dir)
        if fname[0]:
            f = open(fname[0], 'r')
            with f:
                print(f.name)
                self.automatedFundamentalAnalysis(f.name)


    def automatedFundamentalAnalysis(self, fileName):
        grading_metrics = {'Valuation' : ['Fwd P/E', 'PEG', 'P/S', 'P/B', 'P/FCF'],
                            'Profitability' : ['Profit M', 'Oper M', 'Gross M', 'ROE', 'ROA'],
                            'Growth' : ['EPS this Y', 'EPS next Y', 'EPS next 5Y', 'Sales Q/Q', 'EPS Q/Q'],
                            'Performance' : ['Perf Month', 'Perf Quart', 'Perf Half', 'Perf Year', 'Perf YTD', 'Volatility M']
                        }

        grade_scores = {'A+' : 4.3, 'A' : 4.0, 'A-' : 3.7, 'B+' : 3.3, 'B' : 3.0, 'B-' : 2.7, 
                        'C+' : 2.3, 'C' : 2.0, 'C-' : 1.7, 'D+' : 1.3, 'D' : 1.0, 'D-' : 0.7, 'F' : 0.0
                    }

        df = pd.read_csv(fileName, index_col = 0)
        sector_data = collections.defaultdict(lambda : collections.defaultdict(dict))
        data_to_add = collections.defaultdict(list)

        def remove_outliers(S, std):    
            s1 = S[~((S-S.mean()).abs() > std * S.std())]
            return s1[~((s1-s1.mean()).abs() > std * s1.std())]

        def get_metric_grade(sector, metric_name, metric_val):
            lessThan = metric_name in ['Fwd P/E', 'PEG', 'P/S', 'P/B', 'P/FCF', 'Volatility M']
            grade_basis = '10Pct' if lessThan else '90Pct'
            start, change = sector_data[sector][metric_name][grade_basis], sector_data[sector][metric_name]['Std']
            
            grade_map = {'A+': 0, 'A': change, 'A-' : change * 2, 'B+' : change * 3, 'B' : change * 4, 
                        'B-' : change * 5, 'C+' : change * 6, 'C' : change * 7, 'C-' : change * 8, 
                        'D+' : change * 9, 'D' : change * 10, 'D-' : change * 11, 'F' : change * 12
                    }
            
            for grade, val in grade_map.items():
                comparison = start + val if lessThan else start - val
                if lessThan and metric_val < comparison:
                    return grade
                if lessThan == False and metric_val > comparison:
                    return grade
            return 'C'

        def get_metric_val(ticker, metric_name):
            try:
                return float(str(df.loc[df['Ticker'] == ticker][metric_name].values[0]).rstrip("%"))
            except:
                return 0

        def get_category_grades(ticker, sector):
            category_grades = {}
            for category in grading_metrics:
                metric_grades = []
                for metric_name in grading_metrics[category]:
                    metric_grades.append(get_metric_grade(sector, metric_name, get_metric_val(ticker, metric_name)))
                category_grades[category] = metric_grades
                
            for category in category_grades:
                score = 0
                for grade in category_grades[category]:
                    score += grade_scores[grade]
                category_grades[category].append(round(score / len(category_grades[category]), 2))
            return category_grades

        def get_stock_rating(category_grades):
            score = 0
            for category in category_grades:
                score += category_grades[category][-1]
            return round(score * 6.2, 2)

        def convert_to_letter_grade(val):
            for grade in grade_scores:
                if val >= grade_scores[grade]:
                    return grade

        sectors = df['Sector'].unique()
        metrics = df.columns[6: -3]
        for sector in sectors:
            rows = df.loc[df['Sector'] == sector]
            for metric in metrics:
                rows[metric] = rows[metric].str.rstrip('%')
                rows[metric] = pd.to_numeric(rows[metric], errors='coerce')
                data = remove_outliers(rows[metric], 2)
                
                sector_data[sector][metric]['Median'] = data.median(skipna=True)
                sector_data[sector][metric]['10Pct'] = data.quantile(0.1)
                sector_data[sector][metric]['90Pct'] = data.quantile(0.9)
                sector_data[sector][metric]['Std'] = np.std(data, axis=0) / 5


        for row in df.iterrows():
            ticker, sector = row[1]['Ticker'], row[1]['Sector']
            category_grades = get_category_grades(ticker, sector)
            stock_rating = get_stock_rating(category_grades)
            
            data_to_add['Overall Rating'].append(stock_rating)
            data_to_add['Valuation Grade'].append(convert_to_letter_grade(category_grades['Valuation'][-1]))
            data_to_add['Profitability Grade'].append(convert_to_letter_grade(category_grades['Profitability'][-1]))
            data_to_add['Growth Grade'].append(convert_to_letter_grade(category_grades['Growth'][-1]))
            data_to_add['Performance Grade'].append(convert_to_letter_grade(category_grades['Performance'][-1]))

        df['Overall Rating'] = data_to_add['Overall Rating']
        df['Valuation Grade'] = data_to_add['Valuation Grade']
        df['Profitability Grade'] = data_to_add['Profitability Grade']
        df['Growth Grade'] = data_to_add['Growth Grade']
        df['Performance Grade'] = data_to_add['Performance Grade']    
        df['Percent Diff'] = (pd.to_numeric(df['Target Price'], errors='coerce') - pd.to_numeric(df['Price'], errors='coerce')) / pd.to_numeric(df['Price'], errors='coerce') * 100
        print(df)
        df.to_csv('result.csv', index = False)


    def downloadFile(self):
        print("Downloaded!!")


def main():
    app = QApplication(sys.argv)
    gui = GUI()
    gui.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
