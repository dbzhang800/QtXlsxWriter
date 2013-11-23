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
#ifndef QXLSX_FORMAT_H
#define QXLSX_FORMAT_H

#include "xlsxglobal.h"
#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QExplicitlySharedDataPointer>

class FormatTest;

QT_BEGIN_NAMESPACE_XLSX

class Styles;
class Worksheet;
class WorksheetPrivate;
class RichStringPrivate;
class SharedStrings;

class FormatPrivate;
class Q_XLSX_EXPORT Format
{
public:
    enum FontScript
    {
        FontScriptNormal,
        FontScriptSuper,
        FontScriptSub
    };

    enum FontUnderline
    {
        FontUnderlineNone,
        FontUnderlineSingle,
        FontUnderlineDouble,
        FontUnderlineSingleAccounting,
        FontUnderlineDoubleAccounting
    };

    enum HorizontalAlignment
    {
        AlignHGeneral,
        AlignLeft,
        AlignHCenter,
        AlignRight,
        AlignHFill,
        AlignHJustify,
        AlignHMerge,
        AlignHDistributed
    };

    enum VerticalAlignment
    {
        AlignTop,
        AlignVCenter,
        AlignBottom,
        AlignVJustify,
        AlignVDistributed
    };

    enum BorderStyle
    {
        BorderNone,
        BorderThin,
        BorderMedium,
        BorderDashed,
        BorderDotted,
        BorderThick,
        BorderDouble,
        BorderHair,
        BorderMediumDashed,
        BorderDashDot,
        BorderMediumDashDot,
        BorderDashDotDot,
        BorderMediumDashDotDot,
        BorderSlantDashDot
    };

    enum DiagonalBorderType
    {
        DiagonalBorderNone,
        DiagonalBorderDown,
        DiagonalBorderUp,
        DiagnoalBorderBoth
    };

    enum FillPattern
    {
        PatternNone,
        PatternSolid,
        PatternMediumGray,
        PatternDarkGray,
        PatternLightGray,
        PatternDarkHorizontal,
        PatternDarkVertical,
        PatternDarkDown,
        PatternDarkUp,
        PatternDarkGrid,
        PatternDarkTrellis,
        PatternLightHorizontal,
        PatternLightVertical,
        PatternLightDown,
        PatternLightUp,
        PatternLightTrellis,
        PatternGray125,
        PatternGray0625,
        PatternLightGrid
    };

    Format();
    Format(const Format &other);
    Format &operator=(const Format &rhs);
    ~Format();

    int numberFormatIndex() const;
    void setNumberFormatIndex(int format);
    QString numberFormat() const;
    void setNumberFormat(const QString &format);
    void setNumberFormat(int id, const QString &format);
    bool isDateTimeFormat() const;

    int fontSize() const;
    void setFontSize(int size);
    bool fontItalic() const;
    void setFontItalic(bool italic);
    bool fontStrikeOut() const;
    void setFontStrikeOut(bool);
    QColor fontColor() const;
    void setFontColor(const QColor &);
    bool fontBold() const;
    void setFontBold(bool bold);
    FontScript fontScript() const;
    void setFontScript(FontScript);
    FontUnderline fontUnderline() const;
    void setFontUnderline(FontUnderline);
    bool fontOutline() const;
    void setFontOutline(bool outline);
    QString fontName() const;
    void setFontName(const QString &);

    HorizontalAlignment horizontalAlignment() const;
    void setHorizontalAlignment(HorizontalAlignment align);
    VerticalAlignment verticalAlignment() const;
    void setVerticalAlignment(VerticalAlignment align);
    bool textWrap() const;
    void setTextWarp(bool textWrap);
    int rotation() const;
    void setRotation(int rotation);
    int indent() const;
    void setIndent(int indent);
    bool shrinkToFit() const;
    void setShrinkToFit(bool shink);

    void setBorderStyle(BorderStyle style);
    void setBorderColor(const QColor &color);
    BorderStyle leftBorderStyle() const;
    void setLeftBorderStyle(BorderStyle style);
    QColor leftBorderColor() const;
    void setLeftBorderColor(const QColor &color);
    BorderStyle rightBorderStyle() const;
    void setRightBorderStyle(BorderStyle style);
    QColor rightBorderColor() const;
    void setRightBorderColor(const QColor &color);
    BorderStyle topBorderStyle() const;
    void setTopBorderStyle(BorderStyle style);
    QColor topBorderColor() const;
    void setTopBorderColor(const QColor &color);
    BorderStyle bottomBorderStyle() const;
    void setBottomBorderStyle(BorderStyle style);
    QColor bottomBorderColor() const;
    void setBottomBorderColor(const QColor &color);
    BorderStyle diagonalBorderStyle() const;
    void setDiagonalBorderStyle(BorderStyle style);
    DiagonalBorderType diagonalBorderType() const;
    void setDiagonalBorderType(DiagonalBorderType style);
    QColor diagonalBorderColor() const;
    void setDiagonalBorderColor(const QColor &color);

    FillPattern fillPattern() const;
    void setFillPattern(FillPattern pattern);
    QColor patternForegroundColor() const;
    void setPatternForegroundColor(const QColor &color);
    QColor patternBackgroundColor() const;
    void setPatternBackgroundColor(const QColor &color);

    bool locked() const;
    void setLocked(bool locked);
    bool hidden() const;
    void setHidden(bool hidden);

    bool isValid() const;
    bool isEmpty() const;

    bool operator == (const Format &format) const;
    bool operator != (const Format &format) const;

    QVariant property(int propertyId) const;
    void setProperty(int propertyId, const QVariant &value);
    void clearProperty(int propertyId);
    bool hasProperty(int propertyId) const;

    bool boolProperty(int propertyId) const;
    int intProperty(int propertyId) const;
    double doubleProperty(int propertyId) const;
    QString stringProperty(int propertyId) const;
    QColor colorProperty(int propertyId) const;

private:
    friend class Styles;
    friend class Worksheet;
    friend class WorksheetPrivate;
    friend class RichStringPrivate;
    friend class SharedStrings;
    friend class ::FormatTest;
    friend Q_XLSX_EXPORT QDebug operator<<(QDebug, const Format &f);

    bool fontIndexValid() const;
    int fontIndex() const;
    void setFontIndex(int index);
    QByteArray fontKey() const;

    bool hasNumFmtData() const;
    bool hasFontData() const;
    bool hasFillData() const;
    bool hasBorderData() const;
    bool hasAlignmentData() const;
    bool hasProtectionData() const;

    bool borderIndexValid() const;
    QByteArray borderKey() const;
    int borderIndex() const;
    void setBorderIndex(int index);

    bool fillIndexValid() const;
    QByteArray fillKey() const;
    int fillIndex() const;
    void setFillIndex(int index);

    QByteArray formatKey() const;
    bool xfIndexValid() const;
    int xfIndex() const;
    void setXfIndex(int index);
    bool isDxfFormat() const;
    bool dxfIndexValid() const;
    int dxfIndex() const;
    void setDxfIndex(int index);

    int theme() const;

    QExplicitlySharedDataPointer<FormatPrivate> d;
};

#ifndef QT_NO_DEBUG_STREAM
Q_XLSX_EXPORT QDebug operator<<(QDebug dbg, const Format &f);
#endif

QT_END_NAMESPACE_XLSX

#endif // QXLSX_FORMAT_H
