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
#include "xmlstreamwriter_p.h"
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
    m_xf_formats = Format::s_xfFormats;
    m_dxf_formats = Format::s_dxfFormats;

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

    writer.writeStartDocument("1.0", true);
    writer.writeStartElement("styleSheet");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

//    writer.writeStartElement("numFmts");
//    writer.writeEndElement();//numFmts

    writeFonts(writer);
    writeFills(writer);
    writeBorders(writer);

    writer.writeStartElement("cellStyleXfs");
    writer.writeAttribute("count", "1");
    writer.writeStartElement("xf");
    writer.writeAttribute("numFmtId", "0");
    writer.writeAttribute("fontId", "0");
    writer.writeAttribute("fillId", "0");
    writer.writeAttribute("borderId", "0");
    writer.writeEndElement();//xf
    writer.writeEndElement();//cellStyleXfs

    writeCellXfs(writer);

    writer.writeStartElement("cellStyles");
    writer.writeAttribute("count", "1");
    writer.writeStartElement("cellStyle");
    writer.writeAttribute("name", "Normal");
    writer.writeAttribute("xfId", "0");
    writer.writeAttribute("builtinId", "0");
    writer.writeEndElement();//cellStyle
    writer.writeEndElement();//cellStyles

    writeDxfs(writer);

    writer.writeStartElement("tableStyles");
    writer.writeAttribute("count", "0");
    writer.writeAttribute("defaultTableStyle", "TableStyleMedium9");
    writer.writeAttribute("defaultPivotStyle", "PivotStyleLight16");
    writer.writeEndElement();//tableStyles

    writer.writeEndElement();//styleSheet
    writer.writeEndDocument();
}

void Styles::writeFonts(XmlStreamWriter &writer)
{

    writer.writeStartElement("fonts");
    writer.writeAttribute("count", QString::number(m_font_count));
    foreach (Format *format, m_xf_formats) {
        if (format->hasFont()) {
            writer.writeStartElement("font");
            if (format->fontBold())
                writer.writeEmptyElement("b");
            if (format->fontItalic())
                writer.writeEmptyElement("i");
            if (format->fontStrikeOut())
                writer.writeEmptyElement("strike");
            if (format->fontOutline())
                writer.writeEmptyElement("outline");
            if (format->fontShadow())
                writer.writeEmptyElement("shadow");
            if (format->fontUnderline() != Format::FontUnderlineNone) {
                writer.writeEmptyElement("u");
                if (format->fontUnderline() == Format::FontUnderlineDouble)
                    writer.writeAttribute("val", "double");
                else if (format->fontUnderline() == Format::FontUnderlineSingleAccounting)
                    writer.writeAttribute("val", "singleAccounting");
                else if (format->fontUnderline() == Format::FontUnderlineDoubleAccounting)
                    writer.writeAttribute("val", "doubleAccounting");
            }
            if (format->fontScript() != Format::FontScriptNormal) {
                writer.writeEmptyElement("vertAligh");
                if (format->fontScript() == Format::FontScriptSuper)
                    writer.writeAttribute("val", "superscript");
                else
                    writer.writeAttribute("val", "subscript");
            }

            if (!format->isDxfFormat()) {
                writer.writeEmptyElement("sz");
                writer.writeAttribute("val", QString::number(format->fontSize()));
            }

            //font color
            if (format->theme()) {
                writer.writeEmptyElement("color");
                writer.writeAttribute("theme", QString::number(format->theme()));
            } else if (format->colorIndexed()) {
                writer.writeEmptyElement("color");
                writer.writeAttribute("indexed", QString::number(format->colorIndexed()));
            } else if (format->fontColor().isValid()) {
                writer.writeEmptyElement("color");
                QString color = format->fontColor().name();
                writer.writeAttribute("rgb", "FF"+color.mid(1));//remove #
            } else if (!format->isDxfFormat()) {
                writer.writeEmptyElement("color");
                writer.writeAttribute("theme", "1");
            }

            if (!format->isDxfFormat()) {
                writer.writeEmptyElement("name");
                writer.writeAttribute("val", format->fontName());
                writer.writeEmptyElement("family");
                writer.writeAttribute("val", QString::number(format->fontFamily()));
                if (format->fontName() == "Calibri") {
                    writer.writeEmptyElement("scheme");
                    writer.writeAttribute("val", format->fontScheme());
                }
            }

            writer.writeEndElement(); //font
        }
    }
    writer.writeEndElement();//fonts
}

void Styles::writeFills(XmlStreamWriter &writer)
{
    writer.writeStartElement("fills");
    writer.writeAttribute("count", QString::number(m_fill_count));
    //wirte two default fill first
    writer.writeStartElement("fill");
    writer.writeEmptyElement("patternFill");
    writer.writeAttribute("patternType", "none");
    writer.writeEndElement();//fill
    writer.writeStartElement("fill");
    writer.writeEmptyElement("patternFill");
    writer.writeAttribute("patternType", "gray125");
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
        patternStrings[Format::PatternNone] = "none";
        patternStrings[Format::PatternSolid] = "solid";
        patternStrings[Format::PatternMediumGray] = "mediumGray";
        patternStrings[Format::PatternDarkGray] = "darkGray";
        patternStrings[Format::PatternLightGray] = "lightGray";
        patternStrings[Format::PatternDarkHorizontal] = "darkHorizontal";
        patternStrings[Format::PatternDarkVertical] = "darkVertical";
        patternStrings[Format::PatternDarkDown] = "darkDown";
        patternStrings[Format::PatternDarkUp] = "darkUp";
        patternStrings[Format::PatternDarkGrid] = "darkGrid";
        patternStrings[Format::PatternDarkTrellis] = "darkTrellis";
        patternStrings[Format::PatternLightHorizontal] = "lightHorizontal";
        patternStrings[Format::PatternLightVertical] = "lightVertical";
        patternStrings[Format::PatternLightDown] = "lightDown";
        patternStrings[Format::PatternLightUp] = "lightUp";
        patternStrings[Format::PatternLightTrellis] = "lightTrellis";
        patternStrings[Format::PatternGray125] = "gray125";
        patternStrings[Format::PatternGray0625] = "gray0625";
    }

    writer.writeStartElement("fill");
    writer.writeStartElement("patternFill");
    writer.writeAttribute("patternType", patternStrings[format->fillPattern()]);
    if (format->patternForegroundColor().isValid()) {
        writer.writeEmptyElement("fgColor");
        writer.writeAttribute("rgb", "FF"+format->patternForegroundColor().name().mid(1));
    }
    if (format->patternBackgroundColor().isValid()) {
        writer.writeEmptyElement("bgColor");
        writer.writeAttribute("rgb", "FF"+format->patternBackgroundColor().name().mid(1));
    }

    writer.writeEndElement();//patternFill
    writer.writeEndElement();//fill
}

void Styles::writeBorders(XmlStreamWriter &writer)
{
    writer.writeStartElement("borders");
    writer.writeAttribute("count", QString::number(m_borders_count));
    foreach (Format *format, m_xf_formats) {
        if (format->hasBorders()) {
            writer.writeStartElement("border");
            if (format->diagonalBorderType() == Format::DiagonalBorderUp) {
                writer.writeAttribute("diagonalUp", "1");
            } else if (format->diagonalBorderType() == Format::DiagonalBorderDown) {
                writer.writeAttribute("diagonalDown", "1");
            } else if (format->DiagnoalBorderBoth) {
                writer.writeAttribute("diagonalUp", "1");
                writer.writeAttribute("diagonalDown", "1");
            }
            writeSubBorder(writer, "left", format->leftBorderStyle(), format->leftBorderColor());
            writeSubBorder(writer, "right", format->rightBorderStyle(), format->rightBorderColor());
            writeSubBorder(writer, "top", format->topBorderStyle(), format->topBorderColor());
            writeSubBorder(writer, "bottom", format->bottomBorderStyle(), format->bottomBorderColor());

            if (!format->isDxfFormat()) {
                writeSubBorder(writer, "diagonal", format->diagonalBorderStyle(), format->diagonalBorderColor());
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
        stylesString[Format::BorderNone] = "none";
        stylesString[Format::BorderThin] = "thin";
        stylesString[Format::BorderMedium] = "medium";
        stylesString[Format::BorderDashed] = "dashed";
        stylesString[Format::BorderDotted] = "dotted";
        stylesString[Format::BorderThick] = "thick";
        stylesString[Format::BorderDouble] = "double";
        stylesString[Format::BorderHair] = "hair";
        stylesString[Format::BorderMediumDashed] = "mediumDashed";
        stylesString[Format::BorderDashDot] = "dashDot";
        stylesString[Format::BorderMediumDashDot] = "mediumDashDot";
        stylesString[Format::BorderDashDotDot] = "dashDotDot";
        stylesString[Format::BorderMediumDashDotDot] = "mediumDashDotDot";
        stylesString[Format::BorderSlantDashDot] = "slantDashDot";
    }

    writer.writeStartElement(type);
    writer.writeAttribute("style", stylesString[style]);
    writer.writeEmptyElement("color");
    if (color.isValid())
        writer.writeAttribute("rgb", "FF"+color.name().mid(1)); //remove #
    else
        writer.writeAttribute("auto", "1");
    writer.writeEndElement();//type
}

void Styles::writeCellXfs(XmlStreamWriter &writer)
{
    writer.writeStartElement("cellXfs");
    writer.writeAttribute("count", QString::number(m_xf_formats.size()));
    foreach (Format *format, m_xf_formats) {
        int num_fmt_id = format->numberFormat();
        int font_id = format->fontIndex();
        int fill_id = format->fillIndex();
        int border_id = format->borderIndex();
        int xf_id = 0;
        writer.writeStartElement("xf");
        writer.writeAttribute("numFmtId", QString::number(num_fmt_id));
        writer.writeAttribute("fontId", QString::number(font_id));
        writer.writeAttribute("fillId", QString::number(fill_id));
        writer.writeAttribute("borderId", QString::number(border_id));
        writer.writeAttribute("xfId", QString::number(xf_id));
        if (format->numberFormat() > 0)
            writer.writeAttribute("applyNumberFormat", "1");
        if (format->fontIndex() > 0)
            writer.writeAttribute("applyFont", "1");
        if (format->borderIndex() > 0)
            writer.writeAttribute("applyBorder", "1");
        if (format->fillIndex() > 0)
            writer.writeAttribute("applyFill", "1");
        if (format->alignmentChanged())
            writer.writeAttribute("applyAlignment", "1");

        if (format->alignmentChanged()) {
            writer.writeEmptyElement("alignment");
            QString alignH = format->horizontalAlignmentString();
            if (!alignH.isEmpty())
                writer.writeAttribute("horizontal", alignH);
            QString alignV = format->verticalAlignmentString();
            if (!alignV.isEmpty())
                writer.writeAttribute("vertical", alignV);
            if (format->indent())
                writer.writeAttribute("indent", QString::number(format->indent()));
            if (format->textWrap())
                writer.writeAttribute("wrapText", "1");
            if (format->shrinkToFit())
                writer.writeAttribute("shrinkToFit", "1");
            if (format->shrinkToFit())
                writer.writeAttribute("shrinkToFit", "1");
        }

        writer.writeEndElement();//xf
    }
    writer.writeEndElement();//cellXfs
}

void Styles::writeDxfs(XmlStreamWriter &writer)
{
    writer.writeStartElement("dxfs");
    writer.writeAttribute("count", QString::number(m_dxf_formats.size()));
    foreach (Format *format, m_dxf_formats) {
        writer.writeStartElement("dxf");
        writer.writeEndElement();//dxf
    }
    writer.writeEndElement(); //dxfs
}

} //namespace QXlsx
