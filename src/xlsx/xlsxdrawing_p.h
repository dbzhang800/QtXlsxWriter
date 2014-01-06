/****************************************************************************
** Copyright (c) 2013-2014 Debao Zhang <hello@debao.me>
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

//
//  W A R N I N G
//  -------------
//
// This file is not part of the Qt Xlsx API.  It exists for the convenience
// of the Qt Xlsx.  This header file may change from
// version to version without notice, or even be removed.
//
// We mean it.
//

#include <QList>
#include <QString>

class QIODevice;
class QXmlStreamWriter;

namespace QXlsx {

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
    void saveToXmlFile(QIODevice *device) const;
    QByteArray saveToXmlData() const;

    bool embedded;
    int orientation;
    QList <XlsxDrawingDimensionData *> dimensionList;

private:
    void writeTwoCellAnchor(QXmlStreamWriter &writer, int index, XlsxDrawingDimensionData *data) const;
    void writeAbsoluteAnchor(QXmlStreamWriter &writer, int index) const;
    void writePos(QXmlStreamWriter &writer, int x, int y) const;
    void writeExt(QXmlStreamWriter &writer, int cx, int cy) const;
    void writeGraphicFrame(QXmlStreamWriter &writer, int index, const QString &name=QString()) const;
    void writePicture(QXmlStreamWriter &writer, int index, double col_abs, double row_abs, int width, int height, const QString &description) const;
};

} // namespace QXlsx

#endif // QXLSX_DRAWING_H
