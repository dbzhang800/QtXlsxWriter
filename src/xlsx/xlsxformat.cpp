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
#include "xlsxformat_p.h"
#include <QDataStream>
#include <QRegularExpression>
#include <QDebug>

QT_BEGIN_NAMESPACE_XLSX

FormatPrivate::FormatPrivate(Format *p) :
    q_ptr(p)
{
    dirty = true;

    is_dxf_fomat = false;
    xf_index = -1;
    dxf_index = -1;
    xf_indexValid = false;
    dxf_indexValid = false;

    theme = 0;
    color_indexed = 0;
}

/*!
 * \class Format
 * \inmodule QtXlsx
 * \brief Providing the methods and properties that are available for formatting cells in Excel.
 */


/*!
 * \internal
 */
Format::Format() :
    d_ptr(new FormatPrivate(this))
{

}

/*!
 * \internal
 */
Format::~Format()
{
    delete d_ptr;
}

/*!
 * Returns the number format identifier.
 */
int Format::numberFormatIndex() const
{
    Q_D(const Format);
    return d->numberData.formatIndex;
}

/*!
 * Set the number format identifier. The \a format
 * must be a valid built-in number format identifier
 * or the identifier of a custom number format.
 */
void Format::setNumberFormatIndex(int format)
{
    Q_D(Format);
    d->dirty = true;
    d->numberData.formatIndex = format;
    d->numberData._valid = true;
}

/*!
 * Returns the number format string.
 * \note for built-in number formats, this may
 * return an empty string.
 */
QString Format::numberFormat() const
{
    Q_D(const Format);
    return d->numberData.formatString;
}

/*!
 * Set number \a format.
 * http://office.microsoft.com/en-001/excel-help/create-a-custom-number-format-HP010342372.aspx
 */
void Format::setNumberFormat(const QString &format)
{
    Q_D(Format);
    if (format.isEmpty())
        return;
    d->dirty = true;
    d->numberData.formatString = format;
    d->numberData._valid = false; //formatIndex must be re-generated
}

/*!
 * Returns whether the number format is probably a dateTime or not
 */
bool Format::isDateTimeFormat() const
{
    Q_D(const Format);
    if (d->numberData._valid && d->numberData.formatString.isEmpty()) {
        int idx = d->numberData.formatIndex;
        //Built in date time number index
        if ((idx >= 15 && idx <= 22) || (idx >= 45 && idx <= 47))
            return true;
    } else {
        //Gauss from the number string
        QString formatCode = d->numberData.formatString;
        formatCode.remove(QRegularExpression(QStringLiteral("\\[(Green|White|Blue|Magenta|Yellow|Cyan|Red)\\]")));
        if (formatCode.contains(QRegularExpression(QStringLiteral("[dmhys]"))))
            return true;
    }
    return false;
}

/*!
 * \internal
 */
bool Format::numFmtIndexValid() const
{
    Q_D(const Format);
    return d->numberData._valid;
}

/*!
 * \internal
 */
void Format::setNumFmt(int index, const QString &string)
{
    Q_D(Format);
    d->numberData.formatIndex = index;
    d->numberData.formatString = string;
    d->numberData._valid = true;
}

/*!
 * Return the size of the font in points.
 */
int Format::fontSize() const
{
    Q_D(const Format);
    return d->fontData.size;
}

/*!
 * Set the \a size of the font in points.
 */
void Format::setFontSize(int size)
{
    Q_D(Format);
    d->fontData.size = size;
    d->fontData._dirty = true;
}

/*!
 * Return whether the font is italic.
 */
bool Format::fontItalic() const
{
    Q_D(const Format);
    return d->fontData.italic;
}

/*!
 * Turn on/off the italic font.
 */
void Format::setFontItalic(bool italic)
{
    Q_D(Format);
    d->fontData.italic = italic;
    d->fontData._dirty = true;
}

/*!
 * Return whether the font is strikeout.
 */
bool Format::fontStrikeOut() const
{
    Q_D(const Format);
    return d->fontData.strikeOut;
}

/*!
 * Turn on/off the strikeOut font.
 */
void Format::setFontStrikeOut(bool strikeOut)
{
    Q_D(Format);
    d->fontData.strikeOut = strikeOut;
    d->fontData._dirty = true;
}

/*!
 * Return the color of the font.
 */
QColor Format::fontColor() const
{
    Q_D(const Format);
    return d->fontData.color;
}

/*!
 * Set the \a color of the font.
 */
void Format::setFontColor(const QColor &color)
{
    Q_D(Format);
    d->fontData.color = color;
    d->fontData._dirty = true;
}

/*!
 * Return whether the font is bold.
 */
bool Format::fontBold() const
{
    Q_D(const Format);
    return d->fontData.bold;
}

/*!
 * Turn on/off the bold font.
 */
void Format::setFontBold(bool bold)
{
    Q_D(Format);
    d->fontData.bold = bold;
    d->fontData._dirty = true;
}

/*!
 * Return the script style of the font.
 */
Format::FontScript Format::fontScript() const
{
    Q_D(const Format);
    return d->fontData.scirpt;
}

/*!
 * Set the script style of the font.
 */
void Format::setFontScript(FontScript script)
{
    Q_D(Format);
    d->fontData.scirpt = script;
    d->fontData._dirty = true;
}

/*!
 * Return the underline style of the font.
 */
Format::FontUnderline Format::fontUnderline() const
{
    Q_D(const Format);
    return d->fontData.underline;
}

/*!
 * Set the underline style of the font.
 */
void Format::setFontUnderline(FontUnderline underline)
{
    Q_D(Format);
    d->fontData.underline = underline;
    d->fontData._dirty = true;
}

/*!
 * Return whether the font is outline.
 */
bool Format::fontOutline() const
{
    Q_D(const Format);
    return d->fontData.outline;
}

/*!
 * Turn on/off the outline font.
 */
void Format::setFontOutline(bool outline)
{
    Q_D(Format);
    d->fontData.outline = outline;
    d->fontData._dirty = true;
}

/*!
 * Return the name of the font.
 */
QString Format::fontName() const
{
    Q_D(const Format);
    return d->fontData.name;
}

/*!
 * Set the name of the font.
 */
void Format::setFontName(const QString &name)
{
    Q_D(Format);
    d->fontData.name = name;
    d->fontData._dirty = true;
}

/*!
 * \internal
 */
bool Format::fontIndexValid() const
{
    Q_D(const Format);
    return d->fontData.indexValid();
}

/*!
 * \internal
 */
int Format::fontIndex() const
{
    Q_D(const Format);
    return d->fontData.index();
}

/*!
 * \internal
 */
void Format::setFontIndex(int index)
{
    Q_D(Format);
    d->fontData.setIndex(index);
}

/*!
 * \internal
 */
int Format::fontFamily() const
{
    Q_D(const Format);
    return d->fontData.family;
}

/*!
 * \internal
 */
bool Format::fontShadow() const
{
    Q_D(const Format);
    return d->fontData.shadow;
}

/*!
 * \internal
 */
QString Format::fontScheme() const
{
    Q_D(const Format);
    return d->fontData.scheme;
}

/* Internal
 */
QByteArray Format::fontKey() const
{
    Q_D(const Format);
    if (d->fontData._dirty)
        const_cast<FormatPrivate*>(d)->dirty = true; //Make sure formatKey() will be re-generated.
    return d->fontData.key();
}

/*!
 * Return the horizontal alignment.
 */
Format::HorizontalAlignment Format::horizontalAlignment() const
{
    Q_D(const Format);
    return d->alignmentData.alignH;
}

/*!
 * Set the horizontal alignment.
 */
void Format::setHorizontalAlignment(HorizontalAlignment align)
{
    Q_D(Format);
    if (d->alignmentData.indent &&(align != AlignHGeneral && align != AlignLeft &&
                              align != AlignRight && align != AlignHDistributed)) {
        d->alignmentData.indent = 0;
    }

    if (d->alignmentData.shinkToFit && (align == AlignHFill || align == AlignHJustify
                                   || align == AlignHDistributed)) {
        d->alignmentData.shinkToFit = false;
    }

    d->alignmentData.alignH = align;
    d->dirty = true;
}

/*!
 * Return the vertical alignment.
 */
Format::VerticalAlignment Format::verticalAlignment() const
{
    Q_D(const Format);
    return d->alignmentData.alignV;
}

/*!
 * Set the vertical alignment.
 */
void Format::setVerticalAlignment(VerticalAlignment align)
{
    Q_D(Format);
    d->alignmentData.alignV = align;
    d->dirty = true;
}

/*!
 * Return whether the cell text is wrapped.
 */
bool Format::textWrap() const
{
    Q_D(const Format);
    return d->alignmentData.wrap;
}

/*!
 * Enable the text wrap
 */
void Format::setTextWarp(bool wrap)
{
    Q_D(Format);
    if (wrap && d->alignmentData.shinkToFit)
        d->alignmentData.shinkToFit = false;

    d->alignmentData.wrap = wrap;
    d->dirty = true;
}

/*!
 * Return the text rotation.
 */
int Format::rotation() const
{
    Q_D(const Format);
    return d->alignmentData.rotation;
}

/*!
 * Set the text roation. Must be in the range [0, 180] or 255.
 */
void Format::setRotation(int rotation)
{
    Q_D(Format);
    d->alignmentData.rotation = rotation;
    d->dirty = true;
}

/*!
 * Return the text indentation level.
 */
int Format::indent() const
{
    Q_D(const Format);
    return d->alignmentData.indent;
}

/*!
 * Set the text indentation level. Must be less than or equal to 15.
 */
void Format::setIndent(int indent)
{
    Q_D(Format);
    if (indent && (d->alignmentData.alignH != AlignHGeneral
                   && d->alignmentData.alignH != AlignLeft
                   && d->alignmentData.alignH != AlignRight
                   && d->alignmentData.alignH != AlignHJustify)) {
        d->alignmentData.alignH = AlignLeft;
    }
    d->alignmentData.indent = indent;
    d->dirty = true;
}

/*!
 * Return whether the cell is shrink to fit.
 */
bool Format::shrinkToFit() const
{
    Q_D(const Format);
    return d->alignmentData.shinkToFit;
}

/*!
 * Turn on/off shrink to fit.
 */
void Format::setShrinkToFit(bool shink)
{
    Q_D(Format);
    if (shink && d->alignmentData.wrap)
        d->alignmentData.wrap = false;
    if (shink && (d->alignmentData.alignH == AlignHFill
                  || d->alignmentData.alignH == AlignHJustify
                  || d->alignmentData.alignH == AlignHDistributed)) {
        d->alignmentData.alignH = AlignLeft;
    }

    d->alignmentData.shinkToFit = shink;
    d->dirty = true;
}

/*!
 * \internal
 */
bool Format::alignmentChanged() const
{
    Q_D(const Format);
    return d->alignmentData.alignH != AlignHGeneral
            || d->alignmentData.alignV != AlignBottom
            || d->alignmentData.indent != 0
            || d->alignmentData.wrap
            || d->alignmentData.rotation != 0
            || d->alignmentData.shinkToFit;
}

QString Format::horizontalAlignmentString() const
{
    Q_D(const Format);
    QString alignH;
    switch (d->alignmentData.alignH) {
    case Format::AlignLeft:
        alignH = QStringLiteral("left");
        break;
    case Format::AlignHCenter:
        alignH = QStringLiteral("center");
        break;
    case Format::AlignRight:
        alignH = QStringLiteral("right");
        break;
    case Format::AlignHFill:
        alignH = QStringLiteral("fill");
        break;
    case Format::AlignHJustify:
        alignH = QStringLiteral("justify");
        break;
    case Format::AlignHMerge:
        alignH = QStringLiteral("centerContinuous");
        break;
    case Format::AlignHDistributed:
        alignH = QStringLiteral("distributed");
        break;
    default:
        break;
    }
    return alignH;
}

QString Format::verticalAlignmentString() const
{
    Q_D(const Format);
    QString align;
    switch (d->alignmentData.alignV) {
    case AlignTop:
        align = QStringLiteral("top");
        break;
    case AlignVCenter:
        align = QStringLiteral("center");
        break;
    case AlignVJustify:
        align = QStringLiteral("justify");
        break;
    case AlignVDistributed:
        align = QStringLiteral("distributed");
        break;
    default:
        break;
    }
    return align;
}

/*!
 * Set the border style.
 */
void Format::setBorderStyle(BorderStyle style)
{
    setLeftBorderStyle(style);
    setRightBorderStyle(style);
    setBottomBorderStyle(style);
    setTopBorderStyle(style);
}

/*!
 * Set the border color.
 */
void Format::setBorderColor(const QColor &color)
{
    setLeftBorderColor(color);
    setRightBorderColor(color);
    setTopBorderColor(color);
    setBottomBorderColor(color);
}

/*!
 * Return the left border style
 */
Format::BorderStyle Format::leftBorderStyle() const
{
    Q_D(const Format);
    return d->borderData.left;
}

/*!
 * Set the left border style
 */
void Format::setLeftBorderStyle(BorderStyle style)
{
    Q_D(Format);
    d->borderData.left = style;
    d->borderData._dirty = true;
}

/*!
 * Return the left border color
 */
QColor Format::leftBorderColor() const
{
    Q_D(const Format);
    return d->borderData.leftColor;
}

void Format::setLeftBorderColor(const QColor &color)
{
    Q_D(Format);
    d->borderData.leftColor = color;
    d->borderData._dirty = true;
}

Format::BorderStyle Format::rightBorderStyle() const
{
    Q_D(const Format);
    return d->borderData.right;
}

void Format::setRightBorderStyle(BorderStyle style)
{
    Q_D(Format);
    d->borderData.right = style;
    d->borderData._dirty = true;
}

QColor Format::rightBorderColor() const
{
    Q_D(const Format);
    return d->borderData.rightColor;
}

void Format::setRightBorderColor(const QColor &color)
{
    Q_D(Format);
    d->borderData.rightColor = color;
    d->borderData._dirty = true;
}

Format::BorderStyle Format::topBorderStyle() const
{
    Q_D(const Format);
    return d->borderData.top;
}

void Format::setTopBorderStyle(BorderStyle style)
{
    Q_D(Format);
    d->borderData.top = style;
    d->borderData._dirty = true;
}

QColor Format::topBorderColor() const
{
    Q_D(const Format);
    return d->borderData.topColor;
}

void Format::setTopBorderColor(const QColor &color)
{
    Q_D(Format);
    d->borderData.topColor = color;
    d->borderData._dirty = true;
}

Format::BorderStyle Format::bottomBorderStyle() const
{
    Q_D(const Format);
    return d->borderData.bottom;
}

void Format::setBottomBorderStyle(BorderStyle style)
{
    Q_D(Format);
    d->borderData.bottom = style;
    d->borderData._dirty = true;
}

QColor Format::bottomBorderColor() const
{
    Q_D(const Format);
    return d->borderData.bottomColor;
}

void Format::setBottomBorderColor(const QColor &color)
{
    Q_D(Format);
    d->borderData.bottomColor = color;
    d->borderData._dirty = true;
}

Format::BorderStyle Format::diagonalBorderStyle() const
{
    Q_D(const Format);
    return d->borderData.diagonal;
}

void Format::setDiagonalBorderStyle(BorderStyle style)
{
    Q_D(Format);
    d->borderData.diagonal = style;
    d->borderData._dirty = true;
}

Format::DiagonalBorderType Format::diagonalBorderType() const
{
    Q_D(const Format);
    return d->borderData.diagonalType;
}

void Format::setDiagonalBorderType(DiagonalBorderType style)
{
    Q_D(Format);
    d->borderData.diagonalType = style;
    d->borderData._dirty = true;
}

QColor Format::diagonalBorderColor() const
{
    Q_D(const Format);
    return d->borderData.diagonalColor;
}

void Format::setDiagonalBorderColor(const QColor &color)
{
    Q_D(Format);
    d->borderData.diagonalColor = color;
    d->borderData._dirty = true;
}

bool Format::borderIndexValid() const
{
    Q_D(const Format);
    return d->borderData.indexValid();
}

int Format::borderIndex() const
{
    Q_D(const Format);
    return d->borderData.index();
}

void Format::setBorderIndex(int index)
{
    Q_D(Format);
    d->borderData.setIndex(index);
}

/* Internal
 */
QByteArray Format::borderKey() const
{
    Q_D(const Format);
    if (d->borderData._dirty)
        const_cast<FormatPrivate*>(d)->dirty = true; //Make sure formatKey() will be re-generated.

    return d->borderData.key();
}

Format::FillPattern Format::fillPattern() const
{
    Q_D(const Format);
    return d->fillData.pattern;
}

void Format::setFillPattern(FillPattern pattern)
{
    Q_D(Format);
    d->fillData.pattern = pattern;
    d->fillData._dirty = true;
}

QColor Format::patternForegroundColor() const
{
    Q_D(const Format);
    return d->fillData.fgColor;
}

void Format::setPatternForegroundColor(const QColor &color)
{
    Q_D(Format);
    if (color.isValid() && d->fillData.pattern == PatternNone)
        d->fillData.pattern = PatternSolid;
    d->fillData.fgColor = color;
    d->fillData._dirty = true;
}

QColor Format::patternBackgroundColor() const
{
    Q_D(const Format);
    return d->fillData.bgColor;
}

void Format::setPatternBackgroundColor(const QColor &color)
{
    Q_D(Format);
    if (color.isValid() && d->fillData.pattern == PatternNone)
        d->fillData.pattern = PatternSolid;
    d->fillData.bgColor = color;
    d->fillData._dirty = true;
}

bool Format::fillIndexValid() const
{
    Q_D(const Format);
    return d->fillData.indexValid();
}

int Format::fillIndex() const
{
    Q_D(const Format);
    return d->fillData.index();
}

void Format::setFillIndex(int index)
{
    Q_D(Format);
    d->fillData.setIndex(index);
}

/* Internal
 */
QByteArray Format::fillKey() const
{
    Q_D(const Format);
    if (d->fillData._dirty)
        const_cast<FormatPrivate*>(d)->dirty = true; //Make sure formatKey() will be re-generated.

    return d->fillData.key();
}

bool Format::hidden() const
{
    Q_D(const Format);
    return d->protectionData.hidden;
}

void Format::setHidden(bool hidden)
{
    Q_D(Format);
    d->protectionData.hidden = hidden;
    d->dirty = true;
}

bool Format::locked() const
{
    Q_D(const Format);
    return d->protectionData.locked;
}

void Format::setLocked(bool locked)
{
    Q_D(Format);
    d->protectionData.locked = locked;
    d->dirty = true;
}

QByteArray Format::formatKey() const
{
    Q_D(const Format);
    if (d->dirty || d->fontData._dirty || d->borderData._dirty || d->fillData._dirty) {
        QByteArray key;
        QDataStream stream(&key, QIODevice::WriteOnly);
        stream<<fontKey()<<borderKey()<<fillKey()
             <<d->numberData.formatIndex
            <<d->alignmentData.alignH<<d->alignmentData.alignV<<d->alignmentData.indent
           <<d->alignmentData.rotation<<d->alignmentData.shinkToFit<<d->alignmentData.wrap
          <<d->protectionData.hidden<<d->protectionData.locked;
        const_cast<FormatPrivate*>(d)->formatKey = key;
        const_cast<FormatPrivate*>(d)->dirty = false;
        const_cast<FormatPrivate*>(d)->xf_indexValid = false;
        const_cast<FormatPrivate*>(d)->dxf_indexValid = false;
    }

    return d->formatKey;
}

void Format::setXfIndex(int index)
{
    Q_D(Format);
    d->xf_index = index;
    d->xf_indexValid = true;
}

int Format::xfIndex() const
{
    Q_D(const Format);
    return d->xf_index;
}

bool Format::xfIndexValid() const
{
    Q_D(const Format);
    return !d->dirty && d->xf_indexValid;
}

void Format::setDxfIndex(int index)
{
    Q_D(Format);
    d->dxf_index = index;
    d->dxf_indexValid = true;
}

int Format::dxfIndex() const
{
    Q_D(const Format);
    return d->dxf_index;
}

bool Format::dxfIndexValid() const
{
    Q_D(const Format);
    return !d->dirty && d->dxf_indexValid;
}

bool Format::operator ==(const Format &format) const
{
    return this->formatKey() == format.formatKey();
}

bool Format::operator !=(const Format &format) const
{
    return this->formatKey() != format.formatKey();
}

bool Format::isDxfFormat() const
{
    Q_D(const Format);
    return d->is_dxf_fomat;
}

int Format::theme() const
{
    Q_D(const Format);
    return d->theme;
}

int Format::colorIndexed() const
{
    Q_D(const Format);
    return d->color_indexed;
}

QT_END_NAMESPACE_XLSX
