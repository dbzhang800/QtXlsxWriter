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

#include <QObject>
#include <QFont>
#include <QColor>

namespace QXlsx {

class Styles;

class Format : public QObject
{
    Q_OBJECT
public:
    void setFont(const QFont &font);
    void setForegroundColor(const QColor &color);
    void setBackgroundColor(const QColor &color);

private:
    friend class Styles;
    explicit Format(QObject *parent = 0);

    bool isDxfFormat() const;

    //num
    int numFormatIndex() const {return m_num_format_index;}

    //fonts
    bool hasFont() const {return m_has_font;}
    int fontIndex() const {return m_font_index;}
    QString fontName() const;
    bool bold() const;
    bool italic() const;
    bool fontStrikout() const;
    bool fontOutline() const;
    bool fontShadow() const;
    bool fontUnderline() const;
    QColor fontColor() const;
    int fontSize() const;
    int fontFamily() const{return m_font_family;}
    int theme() const {return m_theme;}
    int colorIndexed() const {return m_color_indexed;}
    QString fontScheme() const {return m_font_scheme;}
    void setHasFont(bool has) {m_has_font=has;}

    //fills
    bool hasFill() const {return m_has_fill;}
    int fillIndex() const {return m_fill_index;}

    //borders
    bool hasBorders() const {return m_has_borders;}
    void setHasBorder(bool has) {m_has_borders=has;}
    int borderIndex() const {return m_border_index;}

    bool m_is_dxf_fomat;

    int m_xf_index;
    int m_dxf_index;

    int m_num_format_index;

    bool m_has_font;
    int m_font_index;
    QFont m_font;
    int m_font_family;
    QString m_font_scheme;
    QColor m_font_color;
    QColor m_bg_color;
    QColor m_fg_color;
    int m_theme;
    int m_color_indexed;

    bool m_has_fill;
    int m_fill_index;

    bool m_has_borders;
    int m_border_index;
};

} // namespace QXlsx

#endif // QXLSX_FORMAT_H
