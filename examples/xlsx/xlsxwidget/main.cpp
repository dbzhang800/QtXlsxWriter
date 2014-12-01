#include <QtWidgets>
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"
#include "xlsxsheetmodel.h"

using namespace QXlsx;

int main(int argc, char **argv)
{
    QApplication app(argc, argv);

    //![0]
    QString filePath = QFileDialog::getOpenFileName(0, "Open xlsx file", QString(), "*.xlsx");
    if (filePath.isEmpty())
        return -1;
    //![0]

    //![1]
    QTabWidget tabWidget;
    tabWidget.setWindowTitle(filePath + " - Qt Xlsx Demo");
    tabWidget.setTabPosition(QTabWidget::South);
    //![1]

    //![2]
    Document xlsx(filePath);
    foreach (QString sheetName, xlsx.sheetNames()) {
        Worksheet *sheet = dynamic_cast<Worksheet *>(xlsx.sheet(sheetName));
        if (sheet) {
            QTableView *view = new QTableView(&tabWidget);
            view->setModel(new SheetModel(sheet, view));
            foreach (CellRange range, sheet->mergedCells())
                view->setSpan(range.firstRow()-1, range.firstColumn()-1, range.rowCount(), range.columnCount());
            tabWidget.addTab(view, sheetName);
        }
    }
    //![2]

    tabWidget.show();
    return app.exec();
}
