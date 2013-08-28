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

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>

namespace QXlsx {

class Styles;
class Worksheet;

class Format
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
        PatternGray0625
    };

    int numberFormat() const;
    void setNumberFormat(int format);

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

    bool operator == (const Format &format) const;
    bool operator != (const Format &format) const;

private:
    friend class Styles;
    friend class Worksheet;
    Format();

    struct NumberData
    {
        int formatIndex;
    } m_number;

    struct FontData
    {
        int size;
        bool italic;
        bool strikeOut;
        QColor color;
        bool bold;
        FontScript scirpt;
        FontUnderline underline;
        bool outline;
        bool shadow;
        QString name;
        int family;
        int charset;
        QString scheme;
        int condense;
        int extend;

        //helper member
        bool _dirty; //key re-generated is need.
        QByteArray _key;
        bool _redundant;  //same font already used by some other Formats
        int _index; //index in the Font list
    } m_font;
    bool hasFont() const {return !m_font._redundant;}
    void setFontRedundant(bool redundant) {m_font._redundant = redundant;}
    int fontIndex() const {return m_font._index;}
    void setFontIndex(int index) {m_font._index = index;}
    QByteArray fontKey() const;
    int fontFamily() const{return m_font.family;}
    bool fontShadow() const {return m_font.shadow;}
    QString fontScheme() const {return m_font.scheme;}

    struct AlignmentData
    {
        HorizontalAlignment alignH;
        VerticalAlignment alignV;
        bool wrap;
        int rotation;
        int indent;
        bool shinkToFit;
    } m_alignment;

    bool alignmentChanged() const;
    QString horizontalAlignmentString() const;
    QString verticalAlignmentString() const;

    struct BorderData
    {
        BorderStyle left;
        BorderStyle right;
        BorderStyle top;
        BorderStyle bottom;
        BorderStyle diagonal;
        QColor leftColor;
        QColor rightColor;
        QColor topColor;
        QColor bottomColor;
        QColor diagonalColor;
        DiagonalBorderType diagonalType;

        //helper member
        bool _dirty; //key re-generated is need.
        QByteArray _key;
        bool _redundant;  //same border already used by some other Formats
        int _index; //index in the border list
    } m_border;

    QByteArray borderKey() const;
    bool hasBorders() const {return !m_border._redundant;}
    void setBorderRedundant(bool redundant) {m_border._redundant = redundant;}
    int borderIndex() const {return m_border._index;}
    void setBorderIndex(int index) {m_border._index = index;}

    struct FillData {
        FillPattern pattern;
        QColor bgColor;
        QColor fgColor;

        //helper member
        bool _dirty; //key re-generated is need.
        QByteArray _key;
        bool _redundant;  //same border already used by some other Formats
        int _index; //index in the border list
    } m_fill;

    QByteArray fillKey() const;
    bool hasFill() const {return !m_fill._redundant;}
    void setFillRedundant(bool redundant) {m_fill._redundant = redundant;}
    int fillIndex() const {return m_fill._index;}
    void setFillIndex(int index) {m_fill._index = index;}

    struct ProtectionData {
        bool locked;
        bool hidden;
    } m_protection;

    bool m_dirty; //The key re-generation is need.
    QByteArray m_formatKey;
    QByteArray formatKey() const;

    static QList<Format *> s_xfFormats;
    int m_xf_index;
    int xfIndex(bool generateIfNotValid=true); //Generate index when first called.
    void clearExtraInfos();

    bool m_is_dxf_fomat;
    int m_dxf_index;
    static QList<Format *> s_dxfFormats;
    bool isDxfFormat() const;

    int m_theme;
    int m_color_indexed;
    int theme() const {return m_theme;}
    int colorIndexed() const {return m_color_indexed;}

};

} // namespace QXlsx

#endif // QXLSX_FORMAT_H
