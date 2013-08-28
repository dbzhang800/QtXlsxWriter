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
#include "xlsxformat.h"
#include <QDataStream>
#include <QDebug>

namespace QXlsx {

QList<Format *> Format::s_xfFormats;
QList<Format *> Format::s_dxfFormats;

Format::Format()
{
    m_number.formatIndex = 0;

    m_font.bold = false;
    m_font.color = QColor(Qt::black);
    m_font.italic = false;
    m_font.name = "Calibri";
    m_font.scirpt = FontScriptNormal;
    m_font.size = 11;
    m_font.strikeOut = false;
    m_font.underline = FontUnderlineNone;
    m_font.shadow = false;
    m_font.outline = false;
    m_font.family = 2;
    m_font.scheme = "minor";
    m_font.charset = 0;
    m_font.condense = 0;
    m_font.extend = 0;
    m_font._dirty = true;
    m_font._redundant = false;
    m_font._index = -1;

    m_alignment.alignH = AlignHGeneral;
    m_alignment.alignV = AlignBottom;
    m_alignment.wrap = false;
    m_alignment.rotation = 0;
    m_alignment.indent = 0;
    m_alignment.shinkToFit = false;

    m_border.left = BorderNone;
    m_border.right = BorderNone;
    m_border.top = BorderNone;
    m_border.bottom = BorderNone;
    m_border.diagonal = BorderNone;
    m_border.diagonalType = DiagonalBorderNone;
    m_border.leftColor = QColor();
    m_border.rightColor = QColor();
    m_border.topColor = QColor();
    m_border.bottomColor = QColor();
    m_border.diagonalColor = QColor();
    m_border._dirty = true;
    m_border._redundant = false;
    m_border._index = -1;

    m_fill.pattern = PatternNone;
    m_fill.bgColor = QColor();
    m_fill.fgColor = QColor();
    m_fill._dirty = true;
    m_fill._redundant = false;
    m_fill._index = -1;

    m_protection.locked = false;
    m_protection.hidden = false;

    m_dirty = true;

    m_is_dxf_fomat = false;
    m_xf_index = -1;
    m_dxf_index = -1;

    m_theme = 0;
    m_color_indexed = 0;
}

int Format::numberFormat() const
{
    return m_number.formatIndex;
}

void Format::setNumberFormat(int format)
{
    m_dirty = true;
    m_number.formatIndex = format;
}

int Format::fontSize() const
{
    return m_font.size;
}

void Format::setFontSize(int size)
{
    m_font.size = size;
    m_font._dirty = true;
}

bool Format::fontItalic() const
{
    return m_font.italic;
}

void Format::setFontItalic(bool italic)
{
    m_font.italic = italic;
    m_font._dirty = true;
}

bool Format::fontStrikeOut() const
{
    return m_font.strikeOut;
}

void Format::setFontStrikeOut(bool strikeOut)
{
    m_font.strikeOut = strikeOut;
    m_font._dirty = true;
}

QColor Format::fontColor() const
{
    return m_font.color;
}

void Format::setFontColor(const QColor &color)
{
    m_font.color = color;
    m_font._dirty = true;
}

bool Format::fontBold() const
{
    return m_font.bold;
}

void Format::setFontBold(bool bold)
{
    m_font.bold = bold;
    m_font._dirty = true;
}

Format::FontScript Format::fontScript() const
{
    return m_font.scirpt;
}

void Format::setFontScript(FontScript script)
{
    m_font.scirpt = script;
    m_font._dirty = true;
}

Format::FontUnderline Format::fontUnderline() const
{
    return m_font.underline;
}

void Format::setFontUnderline(FontUnderline underline)
{
    m_font.underline = underline;
    m_font._dirty = true;
}

bool Format::fontOutline() const
{
    return m_font.outline;
}

void Format::setFontOutline(bool outline)
{
    m_font.outline = outline;
    m_font._dirty = true;
}

QString Format::fontName() const
{
    return m_font.name;
}

void Format::setFontName(const QString &name)
{
    m_font.name = name;
    m_font._dirty = true;
}

/* Internal
 */
QByteArray Format::fontKey() const
{
    if (m_font._dirty) {
        QByteArray key;
        QDataStream stream(&key, QIODevice::WriteOnly);
        stream<<m_font.bold<<m_font.charset<<m_font.color<<m_font.condense
             <<m_font.extend<<m_font.family<<m_font.italic<<m_font.name
            <<m_font.outline<<m_font.scheme<<m_font.scirpt<<m_font.shadow
           <<m_font.size<<m_font.strikeOut<<m_font.underline;

        const_cast<Format*>(this)->m_font._key = key;
        const_cast<Format*>(this)->m_font._dirty = false;
        const_cast<Format*>(this)->m_dirty = true; //Make sure formatKey() will be re-generated.
    }

    return m_font._key;
}

Format::HorizontalAlignment Format::horizontalAlignment() const
{
    return m_alignment.alignH;
}

void Format::setHorizontalAlignment(HorizontalAlignment align)
{
    if (m_alignment.indent &&(align != AlignHGeneral && align != AlignLeft &&
                              align != AlignRight && align != AlignHDistributed)) {
        m_alignment.indent = 0;
    }

    if (m_alignment.shinkToFit && (align == AlignHFill || align == AlignHJustify
                                   || align == AlignHDistributed)) {
        m_alignment.shinkToFit = false;
    }

    m_alignment.alignH = align;
    m_dirty = true;
}

Format::VerticalAlignment Format::verticalAlignment() const
{
    return m_alignment.alignV;
}

void Format::setVerticalAlignment(VerticalAlignment align)
{
    m_alignment.alignV = align;
    m_dirty = true;
}

bool Format::textWrap() const
{
    return m_alignment.wrap;
}

void Format::setTextWarp(bool wrap)
{
    if (wrap && m_alignment.shinkToFit)
        m_alignment.shinkToFit = false;

    m_alignment.wrap = wrap;
    m_dirty = true;
}

int Format::rotation() const
{
    return m_alignment.rotation;
}

void Format::setRotation(int rotation)
{
    m_alignment.rotation = rotation;
    m_dirty = true;
}

int Format::indent() const
{
    return m_alignment.indent;
}

void Format::setIndent(int indent)
{
    if (indent && (m_alignment.alignH != AlignHGeneral
                   && m_alignment.alignH != AlignLeft
                   && m_alignment.alignH != AlignRight
                   && m_alignment.alignH != AlignHJustify)) {
        m_alignment.alignH = AlignLeft;
    }
    m_alignment.indent = indent;
    m_dirty = true;
}

bool Format::shrinkToFit() const
{
    return m_alignment.shinkToFit;
}

void Format::setShrinkToFit(bool shink)
{
    if (shink && m_alignment.wrap)
        m_alignment.wrap = false;
    if (shink && (m_alignment.alignH == AlignHFill
                  || m_alignment.alignH == AlignHJustify
                  || m_alignment.alignH == AlignHDistributed)) {
        m_alignment.alignH = AlignLeft;
    }

    m_alignment.shinkToFit = shink;
    m_dirty = true;
}

bool Format::alignmentChanged() const
{
    return m_alignment.alignH != AlignHGeneral
            || m_alignment.alignV != AlignBottom
            || m_alignment.indent != 0
            || m_alignment.wrap
            || m_alignment.rotation != 0
            || m_alignment.shinkToFit;
}

QString Format::horizontalAlignmentString() const
{
    QString alignH;
    switch (m_alignment.alignH) {
    case Format::AlignLeft:
        alignH = "left";
        break;
    case Format::AlignHCenter:
        alignH = "center";
        break;
    case Format::AlignRight:
        alignH = "right";
        break;
    case Format::AlignHFill:
        alignH = "fill";
        break;
    case Format::AlignHJustify:
        alignH = "justify";
        break;
    case Format::AlignHMerge:
        alignH = "centerContinuous";
        break;
    case Format::AlignHDistributed:
        alignH = "distributed";
        break;
    default:
        break;
    }
    return alignH;
}

QString Format::verticalAlignmentString() const
{
    QString align;
    switch (m_alignment.alignV) {
    case AlignTop:
        align = "top";
        break;
    case AlignVCenter:
        align = "center";
        break;
    case AlignVJustify:
        align = "justify";
        break;
    case AlignVDistributed:
        align = "distributed";
        break;
    default:
        break;
    }
    return align;
}

void Format::setBorderStyle(BorderStyle style)
{
    setLeftBorderStyle(style);
    setRightBorderStyle(style);
    setBottomBorderStyle(style);
    setTopBorderStyle(style);
}

void Format::setBorderColor(const QColor &color)
{
    setLeftBorderColor(color);
    setRightBorderColor(color);
    setTopBorderColor(color);
    setBottomBorderColor(color);
}

Format::BorderStyle Format::leftBorderStyle() const
{
    return m_border.left;
}

void Format::setLeftBorderStyle(BorderStyle style)
{
    m_border.left = style;
    m_border._dirty = true;
}

QColor Format::leftBorderColor() const
{
    return m_border.leftColor;
}

void Format::setLeftBorderColor(const QColor &color)
{
    m_border.leftColor = color;
    m_border._dirty = true;
}

Format::BorderStyle Format::rightBorderStyle() const
{
    return m_border.right;
}

void Format::setRightBorderStyle(BorderStyle style)
{
    m_border.right = style;
    m_border._dirty = true;
}

QColor Format::rightBorderColor() const
{
    return m_border.rightColor;
}

void Format::setRightBorderColor(const QColor &color)
{
    m_border.rightColor = color;
    m_border._dirty = true;
}

Format::BorderStyle Format::topBorderStyle() const
{
    return m_border.top;
}

void Format::setTopBorderStyle(BorderStyle style)
{
    m_border.top = style;
    m_border._dirty = true;
}

QColor Format::topBorderColor() const
{
    return m_border.topColor;
}

void Format::setTopBorderColor(const QColor &color)
{
    m_border.topColor = color;
    m_border._dirty = true;
}

Format::BorderStyle Format::bottomBorderStyle() const
{
    return m_border.bottom;
}

void Format::setBottomBorderStyle(BorderStyle style)
{
    m_border.bottom = style;
    m_border._dirty = true;
}

QColor Format::bottomBorderColor() const
{
    return m_border.bottomColor;
}

void Format::setBottomBorderColor(const QColor &color)
{
    m_border.bottomColor = color;
    m_border._dirty = true;
}

Format::BorderStyle Format::diagonalBorderStyle() const
{
    return m_border.diagonal;
}

void Format::setDiagonalBorderStyle(BorderStyle style)
{
    m_border.diagonal = style;
    m_border._dirty = true;
}

Format::DiagonalBorderType Format::diagonalBorderType() const
{
    return m_border.diagonalType;
}

void Format::setDiagonalBorderType(DiagonalBorderType style)
{
    m_border.diagonalType = style;
    m_border._dirty = true;
}

QColor Format::diagonalBorderColor() const
{
    return m_border.diagonalColor;
}

void Format::setDiagonalBorderColor(const QColor &color)
{
    m_border.diagonalColor = color;
    m_border._dirty = true;
}


/* Internal
 */
QByteArray Format::borderKey() const
{
    if (m_border._dirty) {
        QByteArray key;
        QDataStream stream(&key, QIODevice::WriteOnly);
        stream<<m_border.bottom<<m_border.bottomColor
             <<m_border.diagonal<<m_border.diagonalColor<<m_border.diagonalType
            <<m_border.left<<m_border.leftColor
           <<m_border.right<<m_border.rightColor
          <<m_border.top<<m_border.topColor;
        const_cast<Format*>(this)->m_border._key = key;
        const_cast<Format*>(this)->m_border._dirty = false;
        const_cast<Format*>(this)->m_dirty = true; //Make sure formatKey() will be re-generated.
    }

    return m_border._key;
}

Format::FillPattern Format::fillPattern() const
{
    return m_fill.pattern;
}

void Format::setFillPattern(FillPattern pattern)
{
    m_fill.pattern = pattern;
    m_fill._dirty = true;
}

QColor Format::patternForegroundColor() const
{
    return m_fill.fgColor;
}

void Format::setPatternForegroundColor(const QColor &color)
{
    if (color.isValid() && m_fill.pattern == PatternNone)
        m_fill.pattern = PatternSolid;
    m_fill.fgColor = color;
    m_fill._dirty = true;
}

QColor Format::patternBackgroundColor() const
{
    return m_fill.bgColor;
}

void Format::setPatternBackgroundColor(const QColor &color)
{
    if (color.isValid() && m_fill.pattern == PatternNone)
        m_fill.pattern = PatternSolid;
    m_fill.bgColor = color;
    m_fill._dirty = true;
}

/* Internal
 */
QByteArray Format::fillKey() const
{
    if (m_fill._dirty) {
        QByteArray key;
        QDataStream stream(&key, QIODevice::WriteOnly);
        stream<<m_fill.bgColor<<m_fill.fgColor<<m_fill.pattern;
        const_cast<Format*>(this)->m_fill._key = key;
        const_cast<Format*>(this)->m_fill._dirty = false;
        const_cast<Format*>(this)->m_dirty = true; //Make sure formatKey() will be re-generated.
    }

    return m_fill._key;
}

bool Format::hidden() const
{
    return m_protection.hidden;
}

void Format::setHidden(bool hidden)
{
    m_protection.hidden = hidden;
    m_dirty = true;
}

bool Format::locked() const
{
    return m_protection.locked;
}

void Format::setLocked(bool locked)
{
    m_protection.locked = locked;
    m_dirty = true;
}

QByteArray Format::formatKey() const
{
    if (m_dirty || m_font._dirty || m_border._dirty || m_fill._dirty) {
        QByteArray key;
        QDataStream stream(&key, QIODevice::WriteOnly);
        stream<<fontKey()<<borderKey()<<fillKey()
             <<m_number.formatIndex
            <<m_alignment.alignH<<m_alignment.alignV<<m_alignment.indent
           <<m_alignment.rotation<<m_alignment.shinkToFit<<m_alignment.wrap
          <<m_protection.hidden<<m_protection.locked;
        const_cast<Format*>(this)->m_formatKey = key;
        const_cast<Format*>(this)->m_dirty = false;
    }

    return m_formatKey;
}

bool Format::operator ==(const Format &format) const
{
    return this->formatKey() == format.formatKey();
}

bool Format::operator !=(const Format &format) const
{
    return this->formatKey() != format.formatKey();
}

/* Internal
 *
 * This function will be called when wirte the cell contents of worksheet to xml files.
 * Depending on the order of the Format used instead of the Format created, we assign a
 * index to it.
 */
int Format::xfIndex(bool generateIfNotValid)
{
    if (m_xf_index == -1 && generateIfNotValid) { //Generate a valid xf_index for this format
        int index = -1;
        for (int i=0; i<s_xfFormats.size(); ++i) {
            if (*s_xfFormats[i] == *this) {
                index = i;
                break;
            }
        }
        if (index != -1) {
            m_xf_index = index;
        } else {
            m_xf_index = s_xfFormats.size();
            s_xfFormats.append(this);
        }
    }
    return m_xf_index;
}

void Format::clearExtraInfos()
{
    m_xf_index = -1;
    m_dxf_index = -1;
    s_xfFormats.clear();
    s_dxfFormats.clear();
}

bool Format::isDxfFormat() const
{
    return m_is_dxf_fomat;
}

} // namespace QXlsx
