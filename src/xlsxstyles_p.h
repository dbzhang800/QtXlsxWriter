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
#ifndef XLSXSTYLES_H
#define XLSXSTYLES_H

#include <QObject>

class QIODevice;

namespace QXlsx {

class Format;
class XmlStreamWriter;

class Styles : public QObject
{
public:
    explicit Styles(QObject *parent=0);
    Format *addFormat();

    void saveToXmlFile(QIODevice *device);

private:
    void writeFonts(XmlStreamWriter &writer);
    void writeFills(XmlStreamWriter &writer);
    void writeBorders(XmlStreamWriter &writer);
    void writeCellXfs(XmlStreamWriter &writer);
    void writeDxfs(XmlStreamWriter &writer);


    QList<Format *> m_formats;
    QList<Format *> m_xf_formats;
    QList<Format *> m_dxf_formats;

    int m_font_count;
    int m_fill_count;
    int m_borders_count;
};

}
#endif // XLSXSTYLES_H
