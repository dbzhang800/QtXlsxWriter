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
#include "xlsxstyles_p.h"
#include "xlsxformat.h"
#include "xlsxxmlwriter_p.h"
#include <QFile>
#include <QMap>
#include <QDataStream>
#include <QDebug>

namespace QXlsx {


Styles::Styles(QObject *parent) :
    QObject(parent)
{
    m_fill_count = 0;
    m_borders_count = 0;
    m_font_count = 0;
}

Styles::~Styles()
{
    qDeleteAll(m_formats);
}

Format *Styles::addFormat()
{
    Format *format = new Format();

    m_formats.append(format);
    return format;
}

/*
 * This function should be called after worksheet written finished,
 * which means the order of the Formats used have been known to us.
 */
void Styles::prepareStyles()
{
    m_xf_formats = Format::xfFormats();
    m_dxf_formats = Format::dxfFormats();

    if (m_xf_formats.isEmpty())
        m_xf_formats.append(this->addFormat());
    //fonts
    QMap<QByteArray, int> fontsKeyCache;
    foreach (Format *format, m_xf_formats) {
       const QByteArray &key = format->fontKey();
       if (fontsKeyCache.contains(key)) {
           //Font has already been used.
           format->setFontIndex(fontsKeyCache[key]);
           format->setFontRedundant(true);
       } else {
           int index = fontsKeyCache.size();
           fontsKeyCache[key] = index;
           format->setFontIndex(index);
           format->setFontRedundant(false);
       }
    }
    m_font_count = fontsKeyCache.size();

    //borders
    QMap<QByteArray, int> bordersKeyCache;
    foreach (Format *format, m_xf_formats) {
       const QByteArray &key = format->borderKey();
       if (bordersKeyCache.contains(key)) {
           //Border has already been used.
           format->setBorderIndex(bordersKeyCache[key]);
           format->setBorderRedundant(true);
       } else {
           int index = bordersKeyCache.size();
           bordersKeyCache[key] = index;
           format->setBorderIndex(index);
           format->setBorderRedundant(false);
       }
    }
    m_borders_count = bordersKeyCache.size();

    //fills
    QMap<QByteArray, int> fillsKeyCache;
    // The user defined fill properties start from 2 since there are 2
    // default fills: patternType="none" and patternType="gray125".
    {
    QByteArray key;
    QDataStream stream(&key, QIODevice::WriteOnly);
    stream<<QColor()<<QColor()<<Format::PatternNone;
    fillsKeyCache[key] = 0;
    }
    {
    QByteArray key;
    QDataStream stream(&key, QIODevice::WriteOnly);
    stream<<QColor()<<QColor()<<Format::PatternGray125;
    fillsKeyCache[key] = 1;
    }
    foreach (Format *format, m_xf_formats) {
       const QByteArray &key = format->fillKey();
       if (fillsKeyCache.contains(key)) {
           //Border has already been used.
           format->setFillIndex(fillsKeyCache[key]);
           format->setFillRedundant(true);
       } else {
           int index = fillsKeyCache.size();
           fillsKeyCache[key] = index;
           format->setFillIndex(index);
           format->setFillRedundant(false);
       }
    }
    m_fill_count = fillsKeyCache.size();
}

void Styles::clearExtraFormatInfo()
{
    foreach (Format *format, m_formats)
        format->clearExtraInfos();
}

void Styles::saveToXmlFile(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("styleSheet"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));

//    writer.writeStartElement(QStringLiteral("numFmts"));
//    writer.writeEndElement();//numFmts

    writeFonts(writer);
    writeFills(writer);
    writeBorders(writer);

    writer.writeStartElement(QStringLiteral("cellStyleXfs"));
    writer.writeAttribute(QStringLiteral("count"), QStringLiteral("1"));
    writer.writeStartElement(QStringLiteral("xf"));
    writer.writeAttribute(QStringLiteral("numFmtId"), QStringLiteral("0"));
    writer.writeAttribute(QStringLiteral("fontId"), QStringLiteral("0"));
    writer.writeAttribute(QStringLiteral("fillId"), QStringLiteral("0"));
    writer.writeAttribute(QStringLiteral("borderId"), QStringLiteral("0"));
    writer.writeEndElement();//xf
    writer.writeEndElement();//cellStyleXfs

    writeCellXfs(writer);

    writer.writeStartElement(QStringLiteral("cellStyles"));
    writer.writeAttribute(QStringLiteral("count"), QStringLiteral("1"));
    writer.writeStartElement(QStringLiteral("cellStyle"));
    writer.writeAttribute(QStringLiteral("name"), QStringLiteral("Normal"));
    writer.writeAttribute(QStringLiteral("xfId"), QStringLiteral("0"));
    writer.writeAttribute(QStringLiteral("builtinId"), QStringLiteral("0"));
    writer.writeEndElement();//cellStyle
    writer.writeEndElement();//cellStyles

    writeDxfs(writer);

    writer.writeStartElement(QStringLiteral("tableStyles"));
    writer.writeAttribute(QStringLiteral("count"), QStringLiteral("0"));
    writer.writeAttribute(QStringLiteral("defaultTableStyle"), QStringLiteral("TableStyleMedium9"));
    writer.writeAttribute(QStringLiteral("defaultPivotStyle"), QStringLiteral("PivotStyleLight16"));
    writer.writeEndElement();//tableStyles

    writer.writeEndElement();//styleSheet
    writer.writeEndDocument();
}

void Styles::writeFonts(XmlStreamWriter &writer)
{

    writer.writeStartElement(QStringLiteral("fonts"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_font_count));
    foreach (Format *format, m_xf_formats) {
        if (format->hasFont()) {
            writer.writeStartElement(QStringLiteral("font"));
            if (format->fontBold())
                writer.writeEmptyElement(QStringLiteral("b"));
            if (format->fontItalic())
                writer.writeEmptyElement(QStringLiteral("i"));
            if (format->fontStrikeOut())
                writer.writeEmptyElement(QStringLiteral("strike"));
            if (format->fontOutline())
                writer.writeEmptyElement(QStringLiteral("outline"));
            if (format->fontShadow())
                writer.writeEmptyElement(QStringLiteral("shadow"));
            if (format->fontUnderline() != Format::FontUnderlineNone) {
                writer.writeEmptyElement(QStringLiteral("u"));
                if (format->fontUnderline() == Format::FontUnderlineDouble)
                    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("double"));
                else if (format->fontUnderline() == Format::FontUnderlineSingleAccounting)
                    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("singleAccounting"));
                else if (format->fontUnderline() == Format::FontUnderlineDoubleAccounting)
                    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("doubleAccounting"));
            }
            if (format->fontScript() != Format::FontScriptNormal) {
                writer.writeEmptyElement(QStringLiteral("vertAligh"));
                if (format->fontScript() == Format::FontScriptSuper)
                    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("superscript"));
                else
                    writer.writeAttribute(QStringLiteral("val"), QStringLiteral("subscript"));
            }

            if (!format->isDxfFormat()) {
                writer.writeEmptyElement(QStringLiteral("sz"));
                writer.writeAttribute(QStringLiteral("val"), QString::number(format->fontSize()));
            }

            //font color
            if (format->theme()) {
                writer.writeEmptyElement(QStringLiteral("color"));
                writer.writeAttribute(QStringLiteral("theme"), QString::number(format->theme()));
            } else if (format->colorIndexed()) {
                writer.writeEmptyElement(QStringLiteral("color"));
                writer.writeAttribute(QStringLiteral("indexed"), QString::number(format->colorIndexed()));
            } else if (format->fontColor().isValid()) {
                writer.writeEmptyElement(QStringLiteral("color"));
                QString color = format->fontColor().name();
                writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+color.mid(1));//remove #
            } else if (!format->isDxfFormat()) {
                writer.writeEmptyElement(QStringLiteral("color"));
                writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("1"));
            }

            if (!format->isDxfFormat()) {
                writer.writeEmptyElement(QStringLiteral("name"));
                writer.writeAttribute(QStringLiteral("val"), format->fontName());
                writer.writeEmptyElement(QStringLiteral("family"));
                writer.writeAttribute(QStringLiteral("val"), QString::number(format->fontFamily()));
                if (format->fontName() == QLatin1String("Calibri")) {
                    writer.writeEmptyElement(QStringLiteral("scheme"));
                    writer.writeAttribute(QStringLiteral("val"), format->fontScheme());
                }
            }

            writer.writeEndElement(); //font
        }
    }
    writer.writeEndElement();//fonts
}

void Styles::writeFills(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("fills"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_fill_count));
    //wirte two default fill first
    writer.writeStartElement(QStringLiteral("fill"));
    writer.writeEmptyElement(QStringLiteral("patternFill"));
    writer.writeAttribute(QStringLiteral("patternType"), QStringLiteral("none"));
    writer.writeEndElement();//fill
    writer.writeStartElement(QStringLiteral("fill"));
    writer.writeEmptyElement(QStringLiteral("patternFill"));
    writer.writeAttribute(QStringLiteral("patternType"), QStringLiteral("gray125"));
    writer.writeEndElement();//fill
    foreach (Format *format, m_xf_formats) {
        if (format->hasFill()) {
            writeFill(writer, format);
        }
    }
    writer.writeEndElement(); //fills
}

void Styles::writeFill(XmlStreamWriter &writer, Format *format)
{
    static QMap<int, QString> patternStrings;
    if (patternStrings.isEmpty()) {
        patternStrings[Format::PatternNone] = QStringLiteral("none");
        patternStrings[Format::PatternSolid] = QStringLiteral("solid");
        patternStrings[Format::PatternMediumGray] = QStringLiteral("mediumGray");
        patternStrings[Format::PatternDarkGray] = QStringLiteral("darkGray");
        patternStrings[Format::PatternLightGray] = QStringLiteral("lightGray");
        patternStrings[Format::PatternDarkHorizontal] = QStringLiteral("darkHorizontal");
        patternStrings[Format::PatternDarkVertical] = QStringLiteral("darkVertical");
        patternStrings[Format::PatternDarkDown] = QStringLiteral("darkDown");
        patternStrings[Format::PatternDarkUp] = QStringLiteral("darkUp");
        patternStrings[Format::PatternDarkGrid] = QStringLiteral("darkGrid");
        patternStrings[Format::PatternDarkTrellis] = QStringLiteral("darkTrellis");
        patternStrings[Format::PatternLightHorizontal] = QStringLiteral("lightHorizontal");
        patternStrings[Format::PatternLightVertical] = QStringLiteral("lightVertical");
        patternStrings[Format::PatternLightDown] = QStringLiteral("lightDown");
        patternStrings[Format::PatternLightUp] = QStringLiteral("lightUp");
        patternStrings[Format::PatternLightTrellis] = QStringLiteral("lightTrellis");
        patternStrings[Format::PatternGray125] = QStringLiteral("gray125");
        patternStrings[Format::PatternGray0625] = QStringLiteral("gray0625");
    }

    writer.writeStartElement(QStringLiteral("fill"));
    writer.writeStartElement(QStringLiteral("patternFill"));
    writer.writeAttribute(QStringLiteral("patternType"), patternStrings[format->fillPattern()]);
    if (format->patternForegroundColor().isValid()) {
        writer.writeEmptyElement(QStringLiteral("fgColor"));
        writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+format->patternForegroundColor().name().mid(1));
    }
    if (format->patternBackgroundColor().isValid()) {
        writer.writeEmptyElement(QStringLiteral("bgColor"));
        writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+format->patternBackgroundColor().name().mid(1));
    }

    writer.writeEndElement();//patternFill
    writer.writeEndElement();//fill
}

void Styles::writeBorders(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("borders"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_borders_count));
    foreach (Format *format, m_xf_formats) {
        if (format->hasBorders()) {
            writer.writeStartElement(QStringLiteral("border"));
            if (format->diagonalBorderType() == Format::DiagonalBorderUp) {
                writer.writeAttribute(QStringLiteral("diagonalUp"), QStringLiteral("1"));
            } else if (format->diagonalBorderType() == Format::DiagonalBorderDown) {
                writer.writeAttribute(QStringLiteral("diagonalDown"), QStringLiteral("1"));
            } else if (format->DiagnoalBorderBoth) {
                writer.writeAttribute(QStringLiteral("diagonalUp"), QStringLiteral("1"));
                writer.writeAttribute(QStringLiteral("diagonalDown"), QStringLiteral("1"));
            }
            writeSubBorder(writer, QStringLiteral("left"), format->leftBorderStyle(), format->leftBorderColor());
            writeSubBorder(writer, QStringLiteral("right"), format->rightBorderStyle(), format->rightBorderColor());
            writeSubBorder(writer, QStringLiteral("top"), format->topBorderStyle(), format->topBorderColor());
            writeSubBorder(writer, QStringLiteral("bottom"), format->bottomBorderStyle(), format->bottomBorderColor());

            if (!format->isDxfFormat()) {
                writeSubBorder(writer, QStringLiteral("diagonal"), format->diagonalBorderStyle(), format->diagonalBorderColor());
            }
            writer.writeEndElement();//border
        }
    }
    writer.writeEndElement();//borders
}

void Styles::writeSubBorder(XmlStreamWriter &writer, const QString &type, int style, const QColor &color)
{
    if (style == Format::BorderNone) {
        writer.writeEmptyElement(type);
        return;
    }

    static QMap<int, QString> stylesString;
    if (stylesString.isEmpty()) {
        stylesString[Format::BorderNone] = QStringLiteral("none");
        stylesString[Format::BorderThin] = QStringLiteral("thin");
        stylesString[Format::BorderMedium] = QStringLiteral("medium");
        stylesString[Format::BorderDashed] = QStringLiteral("dashed");
        stylesString[Format::BorderDotted] = QStringLiteral("dotted");
        stylesString[Format::BorderThick] = QStringLiteral("thick");
        stylesString[Format::BorderDouble] = QStringLiteral("double");
        stylesString[Format::BorderHair] = QStringLiteral("hair");
        stylesString[Format::BorderMediumDashed] = QStringLiteral("mediumDashed");
        stylesString[Format::BorderDashDot] = QStringLiteral("dashDot");
        stylesString[Format::BorderMediumDashDot] = QStringLiteral("mediumDashDot");
        stylesString[Format::BorderDashDotDot] = QStringLiteral("dashDotDot");
        stylesString[Format::BorderMediumDashDotDot] = QStringLiteral("mediumDashDotDot");
        stylesString[Format::BorderSlantDashDot] = QStringLiteral("slantDashDot");
    }

    writer.writeStartElement(type);
    writer.writeAttribute(QStringLiteral("style"), stylesString[style]);
    writer.writeEmptyElement(QStringLiteral("color"));
    if (color.isValid())
        writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+color.name().mid(1)); //remove #
    else
        writer.writeAttribute(QStringLiteral("auto"), QStringLiteral("1"));
    writer.writeEndElement();//type
}

void Styles::writeCellXfs(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("cellXfs"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_xf_formats.size()));
    foreach (Format *format, m_xf_formats) {
        int num_fmt_id = format->numberFormat();
        int font_id = format->fontIndex();
        int fill_id = format->fillIndex();
        int border_id = format->borderIndex();
        int xf_id = 0;
        writer.writeStartElement(QStringLiteral("xf"));
        writer.writeAttribute(QStringLiteral("numFmtId"), QString::number(num_fmt_id));
        writer.writeAttribute(QStringLiteral("fontId"), QString::number(font_id));
        writer.writeAttribute(QStringLiteral("fillId"), QString::number(fill_id));
        writer.writeAttribute(QStringLiteral("borderId"), QString::number(border_id));
        writer.writeAttribute(QStringLiteral("xfId"), QString::number(xf_id));
        if (format->numberFormat() > 0)
            writer.writeAttribute(QStringLiteral("applyNumberFormat"), QStringLiteral("1"));
        if (format->fontIndex() > 0)
            writer.writeAttribute(QStringLiteral("applyFont"), QStringLiteral("1"));
        if (format->borderIndex() > 0)
            writer.writeAttribute(QStringLiteral("applyBorder"), QStringLiteral("1"));
        if (format->fillIndex() > 0)
            writer.writeAttribute(QStringLiteral("applyFill"), QStringLiteral("1"));
        if (format->alignmentChanged())
            writer.writeAttribute(QStringLiteral("applyAlignment"), QStringLiteral("1"));

        if (format->alignmentChanged()) {
            writer.writeEmptyElement(QStringLiteral("alignment"));
            QString alignH = format->horizontalAlignmentString();
            if (!alignH.isEmpty())
                writer.writeAttribute(QStringLiteral("horizontal"), alignH);
            QString alignV = format->verticalAlignmentString();
            if (!alignV.isEmpty())
                writer.writeAttribute(QStringLiteral("vertical"), alignV);
            if (format->indent())
                writer.writeAttribute(QStringLiteral("indent"), QString::number(format->indent()));
            if (format->textWrap())
                writer.writeAttribute(QStringLiteral("wrapText"), QStringLiteral("1"));
            if (format->shrinkToFit())
                writer.writeAttribute(QStringLiteral("shrinkToFit"), QStringLiteral("1"));
            if (format->shrinkToFit())
                writer.writeAttribute(QStringLiteral("shrinkToFit"), QStringLiteral("1"));
        }

        writer.writeEndElement();//xf
    }
    writer.writeEndElement();//cellXfs
}

void Styles::writeDxfs(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("dxfs"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_dxf_formats.size()));
    foreach (Format *format, m_dxf_formats) {
        writer.writeStartElement(QStringLiteral("dxf"));
        writer.writeEndElement();//dxf
    }
    writer.writeEndElement(); //dxfs
}

} //namespace QXlsx
