#include <QtGui>
#include "xlsxdocument.h"

int main(int argc, char** argv)
{
    QGuiApplication(argc, argv);

    QXlsx::Document xlsx;

    QImage image(400, 300, QImage::Format_RGB32);
    image.fill(Qt::green);
    xlsx.insertImage(5, 5, image);

    xlsx.save();

    return 0;
}
