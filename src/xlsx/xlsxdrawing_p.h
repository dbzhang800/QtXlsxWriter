/****************************************************************************
** Copyright (c) 2013 Debao Zhang <hello@debao.me>
** All right reserved.
**
** Permission is hereby granted, free of charge, to any person obtaining
** a copy of this software and associated documentation files (the
** "Software"), to deal in the Software without restriction, including
** without limitation the rights to use, copy, modify, merge, publish,
** distribute, sublicense, and/or sell copies of the Software, and to
** permit persons to whom the Software is furnished to do so, subject to
** the following conditions:
**
** The above copyright notice and this permission notice shall be
** included in all copies or substantial portions of the Software.
**
** THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
** EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
** MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
** NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
** LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
** OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
** WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
**
****************************************************************************/

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
