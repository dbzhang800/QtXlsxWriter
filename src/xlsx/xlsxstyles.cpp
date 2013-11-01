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
#include "xlsxxmlreader_p.h"
#include "xlsxformat_p.h"
#include "xlsxutility_p.h"
#include <QFile>
#include <QMap>
#include <QDataStream>
#include <QDebug>
#include <QBuffer>

namespace QXlsx {

/*
  When loading from existing .xlsx file. we should create a clean styles object.
  otherwise, default formats should be added.
*/
Styles::Styles(bool createEmpty)
{
    //!Fix me. Should the custom num fmt Id starts with 164 or 176 or others??
    m_nextCustomNumFmtId = 176;

    if (!createEmpty) {
        //Add default Format
        addFormat(createFormat());
        //Add another fill format
        QSharedPointer<FillData> fill = QSharedPointer<FillData>(new FillData);
        fill->pattern = Format::PatternGray125;
        m_fillsList.append(fill);
        m_fillsHash[fill->key()] = fill;
    }
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

Format *Styles::xfFormat(int idx) const
{
    if (idx <0 || idx >= m_xf_formatsList.size())
        return 0;

    return m_xf_formatsList[idx];
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
//            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0_);($#,##0)"), 5);
//            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0_);[Red]($#,##0)"), 6);
//            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0.00_);($#,##0.00)"), 7);
//            m_builtinNumFmtsHash.insert(QStringLiteral("($#,##0.00_);[Red]($#,##0.00)"), 8);
            m_builtinNumFmtsHash.insert(QStringLiteral("0%"), 9);
            m_builtinNumFmtsHash.insert(QStringLiteral("0.00%"), 10);
            m_builtinNumFmtsHash.insert(QStringLiteral("0.00E+00"), 11);
            m_builtinNumFmtsHash.insert(QStringLiteral("# ?/?"), 12);
            m_builtinNumFmtsHash.insert(QStringLiteral("# ?\?/??"), 13);// Note: "??/" is a c++ trigraph, so escape one "?"
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
//            m_builtinNumFmtsHash.insert(QStringLiteral("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(_)"), 41);
//            m_builtinNumFmtsHash.insert(QStringLiteral("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(_)"), 42);
//            m_builtinNumFmtsHash.insert(QStringLiteral("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(_)"), 43);
//            m_builtinNumFmtsHash.insert(QStringLiteral("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(_)"), 44);
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
            format->setNumFmt(m_customNumFmtsHash[str]->formatIndex, str);
        } else {
            //Assign a new fmt Id.
            format->setNumFmt(m_nextCustomNumFmtId, str);

            QSharedPointer<NumberData> fmt(new NumberData(format->d_func()->numberData));
            m_customNumFmtIdMap.insert(m_nextCustomNumFmtId, fmt);
            m_customNumFmtsHash.insert(str, fmt);

            m_nextCustomNumFmtId += 1;
        }
    }

    //Font
    if (!format->fontIndexValid()) {
        if (!m_fontsHash.contains(format->fontKey())) {
            QSharedPointer<FontData> font = QSharedPointer<FontData>(new FontData(format->d_func()->fontData));
            font->setIndex(m_fontsList.size()); //Assign proper index
            m_fontsList.append(font);
            m_fontsHash[font->key()] = font;
        }
        format->setFontIndex(m_fontsHash[format->fontKey()]->index());
    }

    //Fill
    if (!format->fillIndexValid()) {
        if (!m_fillsHash.contains(format->fillKey())) {
            QSharedPointer<FillData> fill = QSharedPointer<FillData>(new FillData(format->d_func()->fillData));
            fill->setIndex(m_fillsList.size()); //Assign proper index
            m_fillsList.append(fill);
            m_fillsHash[fill->key()] = fill;
        }
        format->setFillIndex(m_fillsHash[format->fillKey()]->index());
    }

    //Border
    if (!format->borderIndexValid()) {
        if (!m_bordersHash.contains(format->borderKey())) {
            QSharedPointer<BorderData> border = QSharedPointer<BorderData>(new BorderData(format->d_func()->borderData));
            border->setIndex(m_bordersList.size()); //Assign proper index
            m_bordersList.append(border);
            m_bordersHash[border->key()] = border;
        }
        format->setBorderIndex(m_bordersHash[format->borderKey()]->index());
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
    if (m_customNumFmtIdMap.size() == 0)
        return;

    writer.writeStartElement(QStringLiteral("numFmts"));
    writer.writeAttribute(QStringLiteral("count"), QString::number(m_customNumFmtIdMap.count()));

    QMapIterator<int, QSharedPointer<NumberData> > it(m_customNumFmtIdMap);
    while(it.hasNext()) {
        it.next();
        writer.writeEmptyElement(QStringLiteral("numFmt"));
        writer.writeAttribute(QStringLiteral("numFmtId"), QString::number(it.value()->formatIndex));
        writer.writeAttribute(QStringLiteral("formatCode"), it.value()->formatString);
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
            if (format->rotation())
                writer.writeAttribute(QStringLiteral("textRotation"), QString::number(format->rotation()));
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
        Q_UNUSED(format)
        writer.writeEndElement();//dxf
    }
    writer.writeEndElement(); //dxfs
}

bool Styles::readNumFmts(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("numFmts"));
    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();
    for (int i=0; i<count; ++i) {
        reader.readNextStartElement();
        if (reader.name() != QLatin1String("numFmt"))
            return false;
        QXmlStreamAttributes attributes = reader.attributes();
        QSharedPointer<NumberData> fmt (new NumberData);
        fmt->formatIndex = attributes.value(QLatin1String("numFmtId")).toInt();
        fmt->formatString = attributes.value(QLatin1String("formatCode")).toString();
        if (fmt->formatIndex >= m_nextCustomNumFmtId)
            m_nextCustomNumFmtId = fmt->formatIndex + 1;
        m_customNumFmtIdMap.insert(fmt->formatIndex, fmt);
        m_customNumFmtsHash.insert(fmt->formatString, fmt);

        while (!(reader.name() == QLatin1String("numFmt") && reader.tokenType() == QXmlStreamReader::EndElement))
            reader.readNextStartElement();
    }
    return true;
}

bool Styles::readFonts(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("fonts"));
    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();
    for (int i=0; i<count; ++i) {
        reader.readNextStartElement();
        if (reader.name() != QLatin1String("font"))
            return false;
        QSharedPointer<FontData> font(new FontData);
        while((reader.readNextStartElement(),true)) { //read until font endelement.
            if (reader.tokenType() == QXmlStreamReader::StartElement) {
                if (reader.name() == QLatin1String("b")) {
                    font->bold = true;
                } else if (reader.name() == QLatin1String("i")) {
                    font->italic = true;
                } else if (reader.name() == QLatin1String("strike")) {
                    font->strikeOut = true;
                } else if (reader.name() == QLatin1String("outline")) {
                    font->outline = true;
                } else if (reader.name() == QLatin1String("shadow")) {
                    font->shadow = true;
                } else if (reader.name() == QLatin1String("u")) {
                    QXmlStreamAttributes attributes = reader.attributes();
                    QString value = attributes.value(QLatin1String("val")).toString();
                    if (value == QLatin1String("double"))
                        font->underline = Format::FontUnderlineDouble;
                    else if (value == QLatin1String("doubleAccounting"))
                        font->underline = Format::FontUnderlineDoubleAccounting;
                    else if (value == QLatin1String("singleAccounting"))
                        font->underline = Format::FontUnderlineSingleAccounting;
                    else
                        font->underline = Format::FontUnderlineSingle;
                } else if (reader.name() == QLatin1String("vertAligh")) {
                    QXmlStreamAttributes attributes = reader.attributes();
                    QString value = attributes.value(QLatin1String("val")).toString();
                    if (value == QLatin1String("superscript"))
                        font->scirpt = Format::FontScriptSuper;
                    else
                        font->scirpt = Format::FontScriptSub;
                } else if (reader.name() == QLatin1String("sz")) {
                    font->size = reader.attributes().value(QLatin1String("val")).toInt();
                } else if (reader.name() == QLatin1String("color")) {
                    QXmlStreamAttributes attributes = reader.attributes();
                    if (attributes.hasAttribute(QLatin1String("rgb"))) {
                        QString colorString = attributes.value(QLatin1String("rgb")).toString();
                        font->color = fromARGBString(colorString);
                    } else if (attributes.hasAttribute(QLatin1String("indexed"))) {

                    } else if (attributes.hasAttribute(QLatin1String("theme"))) {

                    }
                } else if (reader.name() == QLatin1String("name")) {
                    font->name = reader.attributes().value(QLatin1String("val")).toString();
                } else if (reader.name() == QLatin1String("family")) {
                    font->family = reader.attributes().value(QLatin1String("val")).toInt();
                } else if (reader.name() == QLatin1String("scheme")) {
                    font->scheme = reader.attributes().value(QLatin1String("val")).toString();
                }
            }

            if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("font"))
                break;
        }
        m_fontsList.append(font);
        m_fontsHash.insert(font->key(), font);
        font->setIndex(m_fontsList.size()-1);//first call key(), then setIndex()
    }
    return true;
}

bool Styles::readFills(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("fills"));

    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();
    for (int i=0; i<count; ++i) {
        reader.readNextStartElement();
        if (reader.name() != QLatin1String("fill") || reader.tokenType() != QXmlStreamReader::StartElement)
            return false;
        readFill(reader);
    }
    return true;
}

bool Styles::readFill(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("fill"));

    static QMap<QString, Format::FillPattern> patternValues;
    if (patternValues.isEmpty()) {
        patternValues[QStringLiteral("none")] = Format::PatternNone;
        patternValues[QStringLiteral("solid")] = Format::PatternSolid;
        patternValues[QStringLiteral("mediumGray")] = Format::PatternMediumGray;
        patternValues[QStringLiteral("darkGray")] = Format::PatternDarkGray;
        patternValues[QStringLiteral("lightGray")] = Format::PatternLightGray;
        patternValues[QStringLiteral("darkHorizontal")] = Format::PatternDarkHorizontal;
        patternValues[QStringLiteral("darkVertical")] = Format::PatternDarkVertical;
        patternValues[QStringLiteral("darkDown")] = Format::PatternDarkDown;
        patternValues[QStringLiteral("darkUp")] = Format::PatternDarkUp;
        patternValues[QStringLiteral("darkGrid")] = Format::PatternDarkGrid;
        patternValues[QStringLiteral("darkTrellis")] = Format::PatternDarkTrellis;
        patternValues[QStringLiteral("lightHorizontal")] = Format::PatternLightHorizontal;
        patternValues[QStringLiteral("lightVertical")] = Format::PatternLightVertical;
        patternValues[QStringLiteral("lightDown")] = Format::PatternLightDown;
        patternValues[QStringLiteral("lightUp")] = Format::PatternLightUp;
        patternValues[QStringLiteral("lightTrellis")] = Format::PatternLightTrellis;
        patternValues[QStringLiteral("gray125")] = Format::PatternGray125;
        patternValues[QStringLiteral("gray0625")] = Format::PatternGray0625;
    }

    QSharedPointer<FillData> fill(new FillData);
    while((reader.readNextStartElement(), true)) { //read until fill endelement
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("patternFill")) {
                QXmlStreamAttributes attributes = reader.attributes();
                QString pattern = attributes.value(QLatin1String("patternType")).toString();
                fill->pattern = patternValues.contains(pattern) ? patternValues[pattern] : Format::PatternNone;
            } else if (reader.name() == QLatin1String("fgColor")) {
                QXmlStreamAttributes attributes = reader.attributes();
                if (attributes.hasAttribute(QLatin1String("rgb"))) {
                    QColor c = fromARGBString(attributes.value(QLatin1String("rgb")).toString());
                    if (fill->pattern == Format::PatternSolid)
                        fill->bgColor = c;
                    else
                        fill->fgColor = c;
                }
            } else if (reader.name() == QLatin1String("bgColor")) {
                QXmlStreamAttributes attributes = reader.attributes();
                if (attributes.hasAttribute(QLatin1String("rgb"))) {
                    QColor c = fromARGBString(attributes.value(QLatin1String("rgb")).toString());
                    if (fill->pattern == Format::PatternSolid)
                        fill->fgColor = c;
                    else
                        fill->bgColor = c;
                }
            }
        }

        if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("fill"))
            break;
    }

    m_fillsList.append(fill);
    m_fillsHash.insert(fill->key(), fill);
    fill->setIndex(m_fillsList.size()-1);//first call key(), then setIndex()

    return true;
}

bool Styles::readBorders(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("borders"));

    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();
    for (int i=0; i<count; ++i) {
        reader.readNextStartElement();
        if (reader.name() != QLatin1String("border") || reader.tokenType() != QXmlStreamReader::StartElement)
            return false;
        readBorder(reader);
    }
    return true;
}

bool Styles::readBorder(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("border"));
    QSharedPointer<BorderData> border(new BorderData);

    QXmlStreamAttributes attributes = reader.attributes();
    bool isUp = attributes.hasAttribute(QLatin1String("diagonalUp"));
    bool isDown = attributes.hasAttribute(QLatin1String("diagonalUp"));
    if (isUp && isDown)
        border->diagonalType = Format::DiagnoalBorderBoth;
    else if (isUp)
        border->diagonalType = Format::DiagonalBorderUp;
    else if (isDown)
        border->diagonalType = Format::DiagonalBorderDown;

    while((reader.readNextStartElement(), true)) { //read until border endelement
        if (reader.tokenType() == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("left"))
                readSubBorder(reader, reader.name().toString(), border->left, border->leftColor);
            else if (reader.name() == QLatin1String("right"))
                readSubBorder(reader, reader.name().toString(), border->right, border->rightColor);
            else if (reader.name() == QLatin1String("top"))
                readSubBorder(reader, reader.name().toString(), border->top, border->topColor);
            else if (reader.name() == QLatin1String("bottom"))
                readSubBorder(reader, reader.name().toString(), border->bottom, border->bottomColor);
            else if (reader.name() == QLatin1String("diagonal"))
                readSubBorder(reader, reader.name().toString(), border->diagonal, border->diagonalColor);
        }

        if (reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("border"))
            break;
    }

    m_bordersList.append(border);
    m_bordersHash.insert(border->key(), border);
    border->setIndex(m_bordersList.size()-1);//first call key(), then setIndex()

    return true;
}

bool Styles::readSubBorder(XmlStreamReader &reader, const QString &name, Format::BorderStyle &style, QColor &color)
{
    Q_ASSERT(reader.name() == name);

    static QMap<QString, Format::BorderStyle> stylesStringsMap;
    if (stylesStringsMap.isEmpty()) {
        stylesStringsMap[QStringLiteral("none")] = Format::BorderNone;
        stylesStringsMap[QStringLiteral("thin")] = Format::BorderThin;
        stylesStringsMap[QStringLiteral("medium")] = Format::BorderMedium;
        stylesStringsMap[QStringLiteral("dashed")] = Format::BorderDashed;
        stylesStringsMap[QStringLiteral("dotted")] = Format::BorderDotted;
        stylesStringsMap[QStringLiteral("thick")] = Format::BorderThick;
        stylesStringsMap[QStringLiteral("double")] = Format::BorderDouble;
        stylesStringsMap[QStringLiteral("hair")] = Format::BorderHair;
        stylesStringsMap[QStringLiteral("mediumDashed")] = Format::BorderMediumDashed;
        stylesStringsMap[QStringLiteral("dashDot")] = Format::BorderDashDot;
        stylesStringsMap[QStringLiteral("mediumDashDot")] = Format::BorderMediumDashDot;
        stylesStringsMap[QStringLiteral("dashDotDot")] = Format::BorderDashDotDot;
        stylesStringsMap[QStringLiteral("mediumDashDotDot")] = Format::BorderMediumDashDotDot;
        stylesStringsMap[QStringLiteral("slantDashDot")] = Format::BorderSlantDashDot;
    }

    QXmlStreamAttributes attributes = reader.attributes();
    if (attributes.hasAttribute(QLatin1String("style"))) {
        QString styleString = attributes.value(QLatin1String("style")).toString();
        if (stylesStringsMap.contains(styleString)) {
            //get style
            style = stylesStringsMap[styleString];
            while((reader.readNextStartElement(),true)) {
                if (reader.tokenType() == QXmlStreamReader::StartElement) {
                    if (reader.name() == QLatin1String("color")) {
                        QXmlStreamAttributes colorAttrs = reader.attributes();
                        if (colorAttrs.hasAttribute(QLatin1String("rgb"))) {
                            QString colorString = colorAttrs.value(QLatin1String("rgb")).toString();
                            //get color
                            color = fromARGBString(colorString);
                        }
                    }

                } else if (reader.tokenType() == QXmlStreamReader::EndElement) {
                    if (reader.name() == name)
                        break;
                }
            }
        }
    }

    return true;
}

bool Styles::readCellXfs(XmlStreamReader &reader)
{
    Q_ASSERT(reader.name() == QLatin1String("cellXfs"));
    QXmlStreamAttributes attributes = reader.attributes();
    int count = attributes.value(QLatin1String("count")).toInt();
    for (int i=0; i<count; ++i) {
        reader.readNextStartElement();
        if (reader.name() != QLatin1String("xf"))
            return false;
        Format *format = createFormat();
        QXmlStreamAttributes xfAttrs = reader.attributes();

//        qDebug()<<reader.name()<<reader.tokenString()<<" .........";
//        for (int i=0; i<xfAttrs.size(); ++i)
//            qDebug()<<"... "<<i<<" "<<xfAttrs[i].name()<<xfAttrs[i].value();

        if (xfAttrs.hasAttribute(QLatin1String("applyNumberFormat"))) {
            int numFmtIndex = xfAttrs.value(QLatin1String("numFmtId")).toInt();
            if (!m_customNumFmtIdMap.contains(numFmtIndex))
                format->setNumberFormatIndex(numFmtIndex);
            else
                format->d_func()->numberData = *m_customNumFmtIdMap[numFmtIndex];
        }

        if (xfAttrs.hasAttribute(QLatin1String("applyFont"))) {
            int fontIndex = xfAttrs.value(QLatin1String("fontId")).toInt();
            if (fontIndex >= m_fontsList.size()) {
                qDebug("Error read styles.xml, cellXfs fontId");
            } else {
                format->d_func()->fontData = *m_fontsList[fontIndex];
            }
        }

        if (xfAttrs.hasAttribute(QLatin1String("applyFill"))) {
            int id = xfAttrs.value(QLatin1String("fillId")).toInt();
            if (id >= m_fillsList.size()) {
                qDebug("Error read styles.xml, cellXfs fillId");
            } else {
                format->d_func()->fillData = *m_fillsList[id];
            }
        }

        if (xfAttrs.hasAttribute(QLatin1String("applyBorder"))) {
            int id = xfAttrs.value(QLatin1String("borderId")).toInt();
            if (id >= m_bordersList.size()) {
                qDebug("Error read styles.xml, cellXfs borderId");
            } else {
                format->d_func()->borderData = *m_bordersList[id];
            }
        }

        if (xfAttrs.hasAttribute(QLatin1String("applyAlignment"))) {
            reader.readNextStartElement();
            if (reader.name() == QLatin1String("alignment")) {
                QXmlStreamAttributes alignAttrs = reader.attributes();

                if (alignAttrs.hasAttribute(QLatin1String("horizontal"))) {
                    static QMap<QString, Format::HorizontalAlignment> alignStringMap;
                    if (alignStringMap.isEmpty()) {
                        alignStringMap.insert(QStringLiteral("left"), Format::AlignLeft);
                        alignStringMap.insert(QStringLiteral("center"), Format::AlignHCenter);
                        alignStringMap.insert(QStringLiteral("right"), Format::AlignRight);
                        alignStringMap.insert(QStringLiteral("justify"), Format::AlignHJustify);
                        alignStringMap.insert(QStringLiteral("centerContinuous"), Format::AlignHMerge);
                        alignStringMap.insert(QStringLiteral("distributed"), Format::AlignHDistributed);
                    }
                    QString str = alignAttrs.value(QLatin1String("horizontal")).toString();
                    if (alignStringMap.contains(str))
                        format->setHorizontalAlignment(alignStringMap[str]);
                }

                if (alignAttrs.hasAttribute(QLatin1String("vertical"))) {
                    static QMap<QString, Format::VerticalAlignment> alignStringMap;
                    if (alignStringMap.isEmpty()) {
                        alignStringMap.insert(QStringLiteral("top"), Format::AlignTop);
                        alignStringMap.insert(QStringLiteral("center"), Format::AlignVCenter);
                        alignStringMap.insert(QStringLiteral("justify"), Format::AlignVJustify);
                        alignStringMap.insert(QStringLiteral("distributed"), Format::AlignVDistributed);
                    }
                    QString str = alignAttrs.value(QLatin1String("vertical")).toString();
                    if (alignStringMap.contains(str))
                        format->setVerticalAlignment(alignStringMap[str]);
                }

                if (alignAttrs.hasAttribute(QLatin1String("indent"))) {
                    int indent = alignAttrs.value(QLatin1String("indent")).toInt();
                    format->setIndent(indent);
                }

                if (alignAttrs.hasAttribute(QLatin1String("textRotation"))) {
                    int rotation = alignAttrs.value(QLatin1String("textRotation")).toInt();
                    format->setRotation(rotation);
                }

                if (alignAttrs.hasAttribute(QLatin1String("wrapText")))
                    format->setTextWarp(true);

                if (alignAttrs.hasAttribute(QLatin1String("shrinkToFit")))
                    format->setShrinkToFit(true);
            }
        }

        addFormat(format);

        //Find the endElement of xf
        while (!(reader.tokenType() == QXmlStreamReader::EndElement && reader.name() == QLatin1String("xf")))
            reader.readNextStartElement();
    }

    return true;
}

bool Styles::loadFromXmlFile(QIODevice *device)
{
    XmlStreamReader reader(device);
    while(!reader.atEnd()) {
        QXmlStreamReader::TokenType token = reader.readNext();
        if (token == QXmlStreamReader::StartElement) {
            if (reader.name() == QLatin1String("numFmts")) {
                readNumFmts(reader);
            } else if (reader.name() == QLatin1String("fonts")) {
                readFonts(reader);
            } else if (reader.name() == QLatin1String("fills")) {
                readFills(reader);
            } else if (reader.name() == QLatin1String("borders")) {
                readBorders(reader);
            } else if (reader.name() == QLatin1String("cellStyleXfs")) {

            } else if (reader.name() == QLatin1String("cellXfs")) {
                readCellXfs(reader);
            } else if (reader.name() == QLatin1String("cellStyles")) {

            }
        }

        if (reader.hasError()) {
            qDebug()<<"Error when read style file: "<<reader.errorString();
        }
    }
    return true;
}

bool Styles::loadFromXmlData(const QByteArray &data)
{
    QBuffer buffer;
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);

    return loadFromXmlFile(&buffer);
}

} //namespace QXlsx
