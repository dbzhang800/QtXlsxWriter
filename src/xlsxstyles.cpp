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

namespace QXlsx {


Styles::Styles(QObject *parent) :
    QObject(parent)
{
    m_fill_count = 2; //Starts from 2
    m_borders_count = 1;
    m_font_count = 0;

    //Add the default cell format
    Format *format = addFormat();
    format->setHasBorder(true);
}

Format *Styles::addFormat()
{
    Format *format = new Format();
    format->setXfIndex(m_formats.size());
    m_font_count += 1;

    m_formats.append(format);
    return format;
}

void Styles::saveToXmlFile(QIODevice *device)
{
    //Todo
    m_xf_formats = m_formats;

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
        //:TODO
        }
    }
    writer.writeEndElement(); //fills
}

void Styles::writeBorders(XmlStreamWriter &writer)
{
    writer.writeStartElement("borders");
    writer.writeAttribute("count", QString::number(m_borders_count));
    foreach (Format *format, m_xf_formats) {
        if (format->hasBorders()) {
            writer.writeStartElement("border");
            writer.writeEmptyElement("left");
            writer.writeEmptyElement("right");
            writer.writeEmptyElement("top");
            writer.writeEmptyElement("bottom");
            if (!format->isDxfFormat()) {
                writer.writeEmptyElement("diagonal");
            }
            writer.writeEndElement();//border
        }
    }
    writer.writeEndElement();//borders
}

void Styles::writeCellXfs(XmlStreamWriter &writer)
{
    writer.writeStartElement("cellXfs");
    writer.writeAttribute("count", QString::number(m_xf_formats.size()));
    foreach (Format *format, m_xf_formats) {
        int num_fmt_id = format->numFormatIndex();
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
        if (format->numFormatIndex() > 0)
            writer.writeAttribute("applyNumberFormat", "1");
        if (format->fontIndex() > 0)
            writer.writeAttribute("applyFont", "1");
        if (format->fillIndex() > 0)
            writer.writeAttribute("applyBorder", "1");
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
