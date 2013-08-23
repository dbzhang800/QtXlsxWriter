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

namespace QXlsx {

Format::Format(QObject *parent) :
    QObject(parent)
{
    m_is_dxf_fomat = false;

    m_xf_index = 0;
    m_dxf_index = 0;

    m_num_format_index = 0;

    m_has_font = false;
    m_font_index = 0;
    m_font_family = 2;
    m_font_scheme = "minor";

    m_font.setFamily("Calibri");
    m_font.setPointSize(11);

    m_theme = 0;
    m_color_indexed = 0;

    m_has_fill = false;
    m_fill_index = 0;

    m_has_borders = false;
    m_border_index = false;
}

bool Format::isDxfFormat() const
{
    return m_is_dxf_fomat;
}

void Format::setFont(const QFont &font)
{
    m_font = font;
}

void Format::setForegroundColor(const QColor &color)
{
    m_fg_color = color;
}

void Format::setBackgroundColor(const QColor &color)
{
    m_bg_color = color;
}

QString Format::fontName() const
{
    return m_font.family();
}

bool Format::bold() const
{
    return m_font.weight() == QFont::Bold;
}

bool Format::italic() const
{
    return m_font.italic();
}

bool Format::fontOutline() const
{
    return false;
}

bool Format::fontShadow() const
{
    return false;
}

bool Format::fontStrikout() const
{
    return m_font.strikeOut();
}

bool Format::fontUnderline() const
{
    return m_font.underline();
}

QColor Format::fontColor() const
{
    return m_font_color;
}

int Format::fontSize() const
{
    return m_font.pointSize();
}


} // namespace QXlsx
