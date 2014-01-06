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

#include "xlsxdrawing_p.h"

#include <QXmlStreamWriter>
#include <QXmlStreamReader>
#include <QBuffer>

namespace QXlsx {

Drawing::Drawing()
{
    embedded = false;
    orientation = 0;
}

QByteArray Drawing::saveToXmlData() const
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);
    return data;
}

void Drawing::saveToXmlFile(QIODevice *device) const
{
    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("xdr:wsDr"));
    writer.writeAttribute(QStringLiteral("xmlns:xdr"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"));
    writer.writeAttribute(QStringLiteral("xmlns:a"), QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main"));

    if (embedded) {
        int index = 1;
        foreach (XlsxDrawingDimensionData *dimension, dimensionList) {
            writeTwoCellAnchor(writer, index, dimension);
            index++;
        }
    } else {
        //write the xdr:absoluteAnchor element
        writeAbsoluteAnchor(writer, 1);
    }

    writer.writeEndElement();//xdr:wsDr
    writer.writeEndDocument();
}

void Drawing::writeTwoCellAnchor(QXmlStreamWriter &writer, int index, XlsxDrawingDimensionData *data) const
{
    writer.writeStartElement(QStringLiteral("xdr:twoCellAnchor"));
    if (data->drawing_type == 2)
        writer.writeAttribute(QStringLiteral("editAs"), QStringLiteral("oneCell"));
//    if (shape)
//        writer.writeAttribute(QStringLiteral("editAs"), );

    writer.writeStartElement(QStringLiteral("xdr:from"));
    writer.writeTextElement(QStringLiteral("xdr:col"), QString::number(data->col_from));
    writer.writeTextElement(QStringLiteral("xdr:colOff"), QString::number((int)data->col_from_offset));
    writer.writeTextElement(QStringLiteral("xdr:row"), QString::number(data->row_from));
    writer.writeTextElement(QStringLiteral("xdr:rowOff"), QString::number((int)data->row_from_offset));
    writer.writeEndElement(); //xdr:from

    writer.writeStartElement(QStringLiteral("xdr:to"));
    writer.writeTextElement(QStringLiteral("xdr:col"), QString::number(data->col_to));
    writer.writeTextElement(QStringLiteral("xdr:colOff"), QString::number((int)data->col_to_offset));
    writer.writeTextElement(QStringLiteral("xdr:row"), QString::number(data->row_to));
    writer.writeTextElement(QStringLiteral("xdr:rowOff"), QString::number((int)data->row_to_offset));
    writer.writeEndElement(); //xdr:to

    if (data->drawing_type == 1) {
        //Graphics frame, xdr:graphicFrame
        writeGraphicFrame(writer, index, data->description);
    } else if (data->drawing_type == 2) {
        //Image, xdr:pic
        writePicture(writer, index, data->col_absolute, data->row_absolute, data->width, data->height, data->description);
    } else {
        //Shape, xdr:sp
    }

    writer.writeEmptyElement(QStringLiteral("xdr:clientData"));
    writer.writeEndElement(); //xdr:twoCellAnchor
}

void Drawing::writeAbsoluteAnchor(QXmlStreamWriter &writer, int index) const
{
    writer.writeStartElement(QStringLiteral("xdr:absoluteAnchor"));
    if (orientation == 0) {
        writePos(writer, 0, 0);
        writeExt(writer, 9308969, 6078325);
    } else {
        writePos(writer, 0, -47625);
        writeExt(writer, 6162675, 6124575);
    }

    writeGraphicFrame(writer, index);
    writer.writeEmptyElement(QStringLiteral("xdr:clientData"));

    writer.writeEndElement(); //xdr:absoluteAnchor
}

void Drawing::writePos(QXmlStreamWriter &writer, int x, int y) const
{
    writer.writeEmptyElement(QStringLiteral("xdr:pos"));
    writer.writeAttribute(QStringLiteral("x"), QString::number(x));
    writer.writeAttribute(QStringLiteral("y"), QString::number(y));
}

void Drawing::writeExt(QXmlStreamWriter &writer, int cx, int cy) const
{
    writer.writeStartElement(QStringLiteral("xdr:ext"));
    writer.writeAttribute(QStringLiteral("cx"), QString::number(cx));
    writer.writeAttribute(QStringLiteral("cy"), QString::number(cy));
}

void Drawing::writeGraphicFrame(QXmlStreamWriter &writer, int index, const QString &name) const
{
    writer.writeStartElement(QStringLiteral("xdr:graphicFrame"));
    writer.writeAttribute(QStringLiteral("macro"), QString());

    writer.writeStartElement(QStringLiteral("xdr:nvGraphicFramePr"));
    writer.writeEmptyElement(QStringLiteral("xdr:cNvPr"));
    writer.writeAttribute(QStringLiteral("id"), QString::number(index+1));
    writer.writeAttribute(QStringLiteral("name"), name.isEmpty() ? QStringLiteral("Chart%1").arg(index): name);
    if (embedded) {
        writer.writeEmptyElement(QStringLiteral("xdr:cNvGraphicFramePr"));
    } else {
        writer.writeStartElement(QStringLiteral("xdr:cNvGraphicFramePr"));
        writer.writeEmptyElement(QStringLiteral("a:graphicFrameLocks"));
        writer.writeAttribute(QStringLiteral("noGrp"), QStringLiteral("1"));
        writer.writeEndElement(); //xdr:cNvGraphicFramePr
    }

    writer.writeEndElement();//xdr:nvGraphicFramePr
    writer.writeEndElement(); //xdr:graphicFrame
}

void Drawing::writePicture(QXmlStreamWriter &writer, int index, double col_abs, double row_abs, int width, int height, const QString &description) const
{
    writer.writeStartElement(QStringLiteral("xdr:pic"));

    writer.writeStartElement(QStringLiteral("xdr:nvPicPr"));
    writer.writeEmptyElement(QStringLiteral("xdr:cNvPr"));
    writer.writeAttribute(QStringLiteral("id"), QString::number(index+1));
    writer.writeAttribute(QStringLiteral("name"), QStringLiteral("Picture%1").arg(index));
    if (!description.isEmpty())
        writer.writeAttribute(QStringLiteral("descr"), description);

    writer.writeStartElement(QStringLiteral("xdr:cNvPicPr"));
    writer.writeEmptyElement(QStringLiteral("a:picLocks"));
    writer.writeAttribute(QStringLiteral("noChangeAspect"), QStringLiteral("1"));
    writer.writeEndElement(); //xdr:cNvPicPr

    writer.writeEndElement(); //xdr:nvPicPr

    writer.writeStartElement(QStringLiteral("xdr:blipFill"));
    writer.writeEmptyElement(QStringLiteral("a:blip"));
    writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
    writer.writeAttribute(QStringLiteral("r:embed"), QStringLiteral("rId%1").arg(index));
    writer.writeStartElement(QStringLiteral("a:stretch"));
    writer.writeEmptyElement(QStringLiteral("a:fillRect"));
    writer.writeEndElement(); //a:stretch
    writer.writeEndElement();//xdr:blipFill

    writer.writeStartElement(QStringLiteral("xdr:spPr"));

    writer.writeStartElement(QStringLiteral("a:xfrm"));
    writer.writeEmptyElement(QStringLiteral("a:off"));
    writer.writeAttribute(QStringLiteral("x"), QString::number((int)col_abs));
    writer.writeAttribute(QStringLiteral("y"), QString::number((int)row_abs));
    writer.writeEmptyElement(QStringLiteral("a:ext"));
    writer.writeAttribute(QStringLiteral("cx"), QString::number(width));
    writer.writeAttribute(QStringLiteral("cy"), QString::number(height));
    writer.writeEndElement(); //a:xfrm

    writer.writeStartElement(QStringLiteral("a:prstGeom"));
    writer.writeAttribute(QStringLiteral("prst"), QStringLiteral("rect"));
    writer.writeEmptyElement(QStringLiteral("a:avLst"));
    writer.writeEndElement(); //a:prstGeom

    writer.writeEndElement(); //xdr:spPr

    writer.writeEndElement(); //xdr:pic
}

} // namespace QXlsx
