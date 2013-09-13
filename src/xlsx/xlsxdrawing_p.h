#ifndef QXLSX_DRAWING_H
#define QXLSX_DRAWING_H

#include <QList>
#include <QString>

class QIODevice;

namespace QXlsx {
class XmlStreamWriter;

struct XlsxDrawingDimensionData
{
    int drawing_type;
    int col_from;
    int row_from;
    double col_from_offset;
    double row_from_offset;
    int col_to;
    int row_to;
    double col_to_offset;
    double row_to_offset;
    int col_absolute;
    int row_absolute;
    int width;
    int height;
    QString description;
    int shape;
};

class Drawing
{
public:
    Drawing();
    void saveToXmlFile(QIODevice *device);

    bool embedded;
    int orientation;
    QList <XlsxDrawingDimensionData *> dimensionList;

private:
    void writeTwoCellAnchor(XmlStreamWriter &writer, int index, XlsxDrawingDimensionData *data);
    void writeAbsoluteAnchor(XmlStreamWriter &writer, int index);
    void writePos(XmlStreamWriter &writer, int x, int y);
    void writeExt(XmlStreamWriter &writer, int cx, int cy);
    void writeGraphicFrame(XmlStreamWriter &writer, int index, const QString &name=QString());
    void writePicture(XmlStreamWriter &writer, int index, double col_abs, double row_abs, int width, int height, const QString &description);
};

} // namespace QXlsx

#endif // QXLSX_DRAWING_H
