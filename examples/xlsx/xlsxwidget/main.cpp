#include <QtWidgets>
#include "xlsxdocument.h"
#include "xlsxsheetmodel.h"

using namespace QXlsx;

int main(int argc, char **argv)
{
    QApplication app(argc, argv);

    //![0]
    QString filePath = QFileDialog::getOpenFileName(0, "Open xlsx file", QString(), ".xlsx");
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
    foreach (QString sheetName, xlsx.worksheetNames()) {
        QTableView *view = new QTableView(&tabWidget);
        SheetModel *model = new SheetModel(xlsx.worksheet(sheetName), view);
        view->setModel(model);
        tabWidget.addTab(view, sheetName);
    }
    //![2]

    tabWidget.show();
    return app.exec();
}
