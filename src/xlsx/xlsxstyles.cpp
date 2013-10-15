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
#include "xlsxxmlwriter_p.h"
#include "xlsxformat_p.h"
#include <QFile>
#include <QMap>
#include <QDataStream>
#include <QDebug>
#include <QBuffer>

namespace QXlsx {


Styles::Styles()
{
    //Add default Format
    addFormat(createFormat());
    //Add another fill format
    QSharedPointer<FillData> fill = QSharedPointer<FillData>(new FillData);
    fill->pattern = Format::PatternGray125;
    m_fillsList.append(fill);
    m_fillsHash[fill->_key] = fill;
}

Styles::~Styles()
{
}

Format *Styles::createFormat()
{
    Format *format = new Format();
    m_createdFormatsList.append(QSharedPointer<Format>(format));

    return format;
}

/*
   Assign index to Font/Fill/Border and Format
*/
void Styles::addFormat(Format *format)
{
    if (!format)
        return;

    //numFmt
    if (!format->numFmtIndexValid()) {
        if (m_builtinNumFmtsHash.isEmpty()) {
            m_builtinNumFmtsHash.insert(QStringLiteral("General"), 0);
            m_builtinNumFmtsHash.insert(QStringLiteral("0"), 1);
            m_builtinNumFmtsHash.insert(QStringLiteral("0.00"), 2);
            m_builtinNumFmtsHash.insert(QStringLiteral("#,##0"), 3);
            m_builtinNumFmtsHash.insert(QStringLiteral("#,##0.00"), 4);
            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0_);($#,##0)"), 5);
            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0_);[Red]($#,##0)"), 6);
            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0.00_);($#,##0.00)"), 7);
            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0.00_);[Red]($#,##0.00)"), 8);
            m_builtinNumFmtsHash.insert(QStringLiteral("0%"), 9);
            m_builtinNumFmtsHash.insert(QStringLiteral("0.00%"), 10);
            m_builtinNumFmtsHash.insert(QStringLiteral("0.00E+00"), 11);
            m_builtinNumFmtsHash.insert(QStringLiteral("# ?/?"), 12);
            m_builtinNumFmtsHash.insert(QStringLiteral("# ??/??"), 13);
            m_builtinNumFmtsHash.insert(QStringLiteral("m/d/yy"), 14);
            m_builtinNumFmtsHash.insert(QStringLiteral("d-mmm-yy"), 15);
            m_builtinNumFmtsHash.insert(QStringLiteral("d-mmm"), 16);
            m_builtinNumFmtsHash.insert(QStringLiteral("mmm-yy"), 17);
            m_builtinNumFmtsHash.insert(QStringLiteral("h:mm AM/PM"), 18);
            m_builtinNumFmtsHash.insert(QStringLiteral("h:mm:ss AM/PM"), 19);
            m_builtinNumFmtsHash.insert(QStringLiteral("h:mm"), 20);
            m_builtinNumFmtsHash.insert(QStringLiteral("h:mm:ss"), 21);
            m_builtinNumFmtsHash.insert(QStringLiteral("m/d/yy h:mm"), 22);

            m_builtinNumFmtsHash.insert(QStringLiteral("(#,##0_);(#,##0)"), 37);
            m_builtinNumFmtsHash.insert(QStringLiteral("(#,##0_);[Red](#,##0)"), 38);
            m_builtinNumFmtsHash.insert(QStringLiteral("(#,##0.00_);(#,##0.00)"), 39);
            m_builtinNumFmtsHash.insert(QStringLiteral("(#,##0.00_);[Red](#,##0.00)"), 40);
            m_builtinNumFmtsHash.insert(QStringLiteral("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(_)"), 41);
            m_builtinNumFmtsHash.insert(QStringLiteral("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(_)"), 42);
            m_builtinNumFmtsHash.insert(QStringLiteral("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(_)"), 43);
            m_builtinNumFmtsHash.insert(QStringLiteral("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(_)"), 44);
            m_builtinNumFmtsHash.insert(QStringLiteral("mm:ss"), 45);
            m_builtinNumFmtsHash.insert(QStringLiteral("[h]:mm:ss"), 46);
            m_builtinNumFmtsHash.insert(QStringLiteral("mm:ss.0"), 47);
            m_builtinNumFmtsHash.insert(QStringLiteral("##0.0E+0"), 48);
            m_builtinNumFmtsHash.insert(QStringLiteral("@"), 49);
        }
        const QString str = format->numberFormat();
        //Assign proper number format index
        if (m_builtinNumFmtsHash.contains(str)) {
            format->setNumFmt(m_builtinNumFmtsHash[str], str);
        } else if (m_customNumFmtsHash.contains(str)) {
            format->setNumFmt(m_customNumFmtsHash[str], str);
        } else {
            int idx = 164 + m_customNumFmts.size();
            m_customNumFmts.append(str);
            m_customNumFmtsHash.insert(str, idx);
            format->setNumFmt(idx, str);
        }
    }

    //Font
    if (!format->fontIndexValid()) {
        if (!m_fontsHash.contains(format->fontKey())) {
            QSharedPointer<FontData> font = QSharedPointer<FontData>(new FontData(format->d_func()->fontData));
            font->_index = m_fontsList.size(); //Assign proper index
            m_fontsList.append(font);
            m_fontsHash[font->_key] = font;
        }
        format->setFontIndex(m_fontsHash[format->fontKey()]->_index);
    }

    //Fill
    if (!format->fillIndexValid()) {
        if (!m_fillsHash.contains(format->fillKey())) {
            QSharedPointer<FillData> fill = QSharedPointer<FillData>(new FillData(format->d_func()->fillData));
            fill->_index = m_fillsList.size(); //Assign proper index
            m_fillsList.append(fill);
            m_fillsHash[fill->_key] = fill;
        }
        format->setFillIndex(m_fillsHash[format->fillKey()]->_index);
    }

    //Border
    if (!format->borderIndexValid()) {
        if (!m_bordersHash.contains(format->borderKey())) {
            QSharedPointer<BorderData> border = QSharedPointer<BorderData>(new BorderData(format->d_func()->borderData));
            border->_index = m_bordersList.size(); //Assign proper index
            m_bordersList.append(border);
            m_bordersHash[border->_key] = border;
        }
        format->setBorderIndex(m_bordersHash[format->borderKey()]->_index);
    }

    //Format
    if (format->isDxfFormat()) {
        if (!format->dxfIndexValid()) {
            if (!m_dxf_formatsHash.contains(format->formatKey())) {
                format->setDxfIndex(m_dxf_formatsList.size());
                m_dxf_formatsList.append(format);
                m_dxf_formatsHash[format->formatKey()] = format;
            } else {
                format->setDxfIndex(m_dxf_formatsHash[format->formatKey()]->dxfIndex());
            }
        }
    } else {
        if (!format->xfIndexValid()) {
            if (!m_xf_formatsHash.contains(format->formatKey())) {
                format->setXfIndex(m_xf_formatsList.size());
                m_xf_formatsList.append(format);
                m_xf_formatsHash[format->formatKey()] = format;
            } else {
                format->setXfIndex(m_xf_formatsHash[format->formatKey()]->xfIndex());
            }
        }
    }
}

QByteArray Styles::saveToXmlData()
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);

    return data;
}

void Styles::saveToXmlFile(QIODevice *device)
{
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("styleSheet"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));

    writeNumFmts(writer);
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

void Styles::writeNumFmts(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("numFmts"));
    for (int i=0; i<m_customNumFmts.size(); ++i) {
        writer.writeEmptyElement(QStringLiteral("numFmt"));
        writer.writeAttribute(QStringLiteral("numFmtId"), QString::number(164 + i));
        writer.writeAttribute(QStringLiteral("formatCode"), m_customNumFmts[i]);
    }
    writer.writeEndElement();//numFmts
}

/*
 not consider dxf format.
*/
void Styles::writeFonts(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("fonts"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_fontsList.count()));
    for (int i=0; i<m_fontsList.size(); ++i) {
        QSharedPointer<FontData> font = m_fontsList[i];

        writer.writeStartElement(QStringLiteral("font"));
        if (font->bold)
            writer.writeEmptyElement(QStringLiteral("b"));
        if (font->italic)
            writer.writeEmptyElement(QStringLiteral("i"));
        if (font->strikeOut)
            writer.writeEmptyElement(QStringLiteral("strike"));
        if (font->outline)
            writer.writeEmptyElement(QStringLiteral("outline"));
        if (font->shadow)
            writer.writeEmptyElement(QStringLiteral("shadow"));
        if (font->underline != Format::FontUnderlineNone) {
            writer.writeEmptyElement(QStringLiteral("u"));
            if (font->underline == Format::FontUnderlineDouble)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("double"));
            else if (font->underline == Format::FontUnderlineSingleAccounting)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("singleAccounting"));
            else if (font->underline == Format::FontUnderlineDoubleAccounting)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("doubleAccounting"));
        }
        if (font->scirpt != Format::FontScriptNormal) {
            writer.writeEmptyElement(QStringLiteral("vertAligh"));
            if (font->scirpt == Format::FontScriptSuper)
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("superscript"));
            else
                writer.writeAttribute(QStringLiteral("val"), QStringLiteral("subscript"));
        }

        writer.writeEmptyElement(QStringLiteral("sz"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(font->size));

        if (font->color.isValid()) {
            writer.writeEmptyElement(QStringLiteral("color"));
            QString color = font->color.name();
            writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+color.mid(1));//remove #
        }

        writer.writeEmptyElement(QStringLiteral("name"));
        writer.writeAttribute(QStringLiteral("val"), font->name);
        writer.writeEmptyElement(QStringLiteral("family"));
        writer.writeAttribute(QStringLiteral("val"), QString::number(font->family));
        if (font->name == QLatin1String("Calibri")) {
            writer.writeEmptyElement(QStringLiteral("scheme"));
            writer.writeAttribute(QStringLiteral("val"), font->scheme);
        }

//        if (!format->isDxfFormat()) {
//            writer.writeEmptyElement(QStringLiteral("sz"));
//            writer.writeAttribute(QStringLiteral("val"), QString::number(format->fontSize()));
//        }
//
//        //font color
//        if (format->theme()) {
//            writer.writeEmptyElement(QStringLiteral("color"));
//            writer.writeAttribute(QStringLiteral("theme"), QString::number(format->theme()));
//        } else if (format->colorIndexed()) {
//            writer.writeEmptyElement(QStringLiteral("color"));
//            writer.writeAttribute(QStringLiteral("indexed"), QString::number(format->colorIndexed()));
//        } else if (format->fontColor().isValid()) {
//            writer.writeEmptyElement(QStringLiteral("color"));
//            QString color = format->fontColor().name();
//            writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+color.mid(1));//remove #
//        } else if (!format->isDxfFormat()) {
//            writer.writeEmptyElement(QStringLiteral("color"));
//            writer.writeAttribute(QStringLiteral("theme"), QStringLiteral("1"));
//        }

//        if (!format->isDxfFormat()) {
//            writer.writeEmptyElement(QStringLiteral("name"));
//            writer.writeAttribute(QStringLiteral("val"), format->fontName());
//            writer.writeEmptyElement(QStringLiteral("family"));
//            writer.writeAttribute(QStringLiteral("val"), QString::number(format->fontFamily()));
//            if (format->fontName() == QLatin1String("Calibri")) {
//                writer.writeEmptyElement(QStringLiteral("scheme"));
//                writer.writeAttribute(QStringLiteral("val"), format->fontScheme());
//            }
//        }
        writer.writeEndElement(); //font
    }
    writer.writeEndElement();//fonts
}

void Styles::writeFills(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("fills"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_fillsList.size()));

    for (int i=0; i<m_fillsList.size(); ++i) {
        QSharedPointer<FillData> fill = m_fillsList[i];
        writeFill(writer, fill.data());
    }
    writer.writeEndElement(); //fills
}

void Styles::writeFill(XmlStreamWriter &writer, FillData *fill)
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
    writer.writeAttribute(QStringLiteral("patternType"), patternStrings[fill->pattern]);
    // For a solid fill, Excel reverses the role of foreground and background colours
    if (fill->fgColor.isValid()) {
        writer.writeEmptyElement(fill->pattern == Format::PatternSolid ? QStringLiteral("bgColor") : QStringLiteral("fgColor"));
        writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+fill->fgColor.name().mid(1));
    }
    if (fill->bgColor.isValid()) {
        writer.writeEmptyElement(fill->pattern == Format::PatternSolid ? QStringLiteral("fgColor") : QStringLiteral("bgColor"));
        writer.writeAttribute(QStringLiteral("rgb"), QStringLiteral("FF")+fill->bgColor.name().mid(1));
    }

    writer.writeEndElement();//patternFill
    writer.writeEndElement();//fill
}

void Styles::writeBorders(XmlStreamWriter &writer)
{
    writer.writeStartElement(QStringLiteral("borders"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_bordersList.count()));
    for (int i=0; i<m_bordersList.size(); ++i) {
        QSharedPointer<BorderData> border = m_bordersList[i];

        writer.writeStartElement(QStringLiteral("border"));
        if (border->diagonalType == Format::DiagonalBorderUp) {
            writer.writeAttribute(QStringLiteral("diagonalUp"), QStringLiteral("1"));
        } else if (border->diagonalType == Format::DiagonalBorderDown) {
            writer.writeAttribute(QStringLiteral("diagonalDown"), QStringLiteral("1"));
        } else if (border->diagonalType == Format::DiagnoalBorderBoth) {
            writer.writeAttribute(QStringLiteral("diagonalUp"), QStringLiteral("1"));
            writer.writeAttribute(QStringLiteral("diagonalDown"), QStringLiteral("1"));
        }
        writeSubBorder(writer, QStringLiteral("left"), border->left, border->leftColor);
        writeSubBorder(writer, QStringLiteral("right"), border->right, border->rightColor);
        writeSubBorder(writer, QStringLiteral("top"), border->top, border->topColor);
        writeSubBorder(writer, QStringLiteral("bottom"), border->bottom, border->bottomColor);

//        if (!format->isDxfFormat()) {
        writeSubBorder(writer, QStringLiteral("diagonal"), border->diagonal, border->diagonalColor);
//        }
        writer.writeEndElement();//border
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
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_xf_formatsList.size()));
    foreach (Format *format, m_xf_formatsList) {
        int num_fmt_id = format->numberFormatIndex();
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
        if (format->numberFormatIndex() > 0)
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
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_dxf_formatsList.size()));
    foreach (Format *format, m_dxf_formatsList) {
        writer.writeStartElement(QStringLiteral("dxf"));
        writer.writeEndElement();//dxf
    }
    writer.writeEndElement(); //dxfs
}

QSharedPointer<Styles> Styles::loadFromXmlFile(QIODevice *device)
{

    return QSharedPointer<Styles>(new Styles);
}

QSharedPointer<Styles> Styles::loadFromXmlData(const QByteArray &data)
{
    QBuffer buffer;
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);

    return loadFromXmlFile(&buffer);
}

} //namespace QXlsx
