#include <QtGui>
#include "xlsxdocument.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main(int argc, char** argv)
{
    QGuiApplication(argc, argv);

    QXlsx::Document xlsx;

    QImage image(400, 300, QImage::Format_RGB32);
    image.fill(Qt::green);
    xlsx.insertImage(5, 5, image);

    xlsx.saveAs(DATA_PATH"Test.xlsx");

    return 0;
}
