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
#ifndef XLSXFORMAT_P_H
#define XLSXFORMAT_P_H
#include "xlsxformat.h"

namespace QXlsx {

struct NumberData
{
    NumberData() : formatIndex(0), _valid(true) {}

    int formatIndex;
    QString formatString;

    bool _valid;
};

struct FontData
{
    FontData() :
        size(11), italic(false), strikeOut(false), color(QColor()), bold(false)
      , scirpt(Format::FontScriptNormal), underline(Format::FontUnderlineNone)
      , outline(false), shadow(false), name(QStringLiteral("Calibri")), family(2), charset(0)
      , scheme(QStringLiteral("minor")), condense(0), extend(0)
      , _dirty(true), _indexValid(false), _index(-1)

    {}

    int size;
    bool italic;
    bool strikeOut;
    QColor color;
    bool bold;
    Format::FontScript scirpt;
    Format::FontUnderline underline;
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
    bool _indexValid;  //has a valid index, so no need to assign a new one
    int _index; //index in the Font list
};

struct AlignmentData
{
    AlignmentData() :
        alignH(Format::AlignHGeneral), alignV(Format::AlignBottom)
      , wrap(false), rotation(0), indent(0), shinkToFit(false)
    {}

    Format::HorizontalAlignment alignH;
    Format::VerticalAlignment alignV;
    bool wrap;
    int rotation;
    int indent;
    bool shinkToFit;
};

struct BorderData
{
    BorderData() :
        left(Format::BorderNone), right(Format::BorderNone), top(Format::BorderNone)
      ,bottom(Format::BorderNone), diagonal(Format::BorderNone)
      ,diagonalType(Format::DiagonalBorderNone)
      ,_dirty(true), _indexValid(false), _index(-1)
    {}

    Format::BorderStyle left;
    Format::BorderStyle right;
    Format::BorderStyle top;
    Format::BorderStyle bottom;
    Format::BorderStyle diagonal;
    QColor leftColor;
    QColor rightColor;
    QColor topColor;
    QColor bottomColor;
    QColor diagonalColor;
    Format::DiagonalBorderType diagonalType;

    //helper member
    bool _dirty; //key re-generated is need.
    QByteArray _key;
    bool _indexValid;  //has a valid index, so no need to assign a new one
    int _index; //index in the border list
};

struct FillData {
    FillData() :
        pattern(Format::PatternNone)
      ,_dirty(true), _indexValid(false), _index(-1)
    {}

    Format::FillPattern pattern;
    QColor bgColor;
    QColor fgColor;

    //helper member
    bool _dirty; //key re-generated is need.
    QByteArray _key;
    bool _indexValid;  //has a valid index, so no need to assign a new one
    int _index; //index in the border list
};

struct ProtectionData {
    ProtectionData() :
        locked(false), hidden(false)
    {}

    bool locked;
    bool hidden;
};

class FormatPrivate
{
    Q_DECLARE_PUBLIC(Format)
public:
    FormatPrivate(Format *p);

    NumberData numberData;
    FontData fontData;
    AlignmentData alignmentData;
    BorderData borderData;
    FillData fillData;
    ProtectionData protectionData;

    bool dirty; //The key re-generation is need.
    QByteArray formatKey;

    static QList<Format *> s_xfFormats;
    int xf_index;
    bool xf_indexValid;

    static QList<Format *> s_dxfFormats;
    bool is_dxf_fomat;
    int dxf_index;
    bool dxf_indexValid;

    int theme;
    int color_indexed;

    Format *q_ptr;
};

}

#endif // XLSXFORMAT_P_H
