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

//
//  W A R N I N G
//  -------------
//
// This file is not part of the Qt Xlsx API.  It exists for the convenience
// of the Qt Xlsx.  This header file may change from
// version to version without notice, or even be removed.
//
// We mean it.
//

#include "xlsxglobal.h"
#include "xlsxformat.h"
#include <QSharedPointer>
#include <QHash>
#include <QList>
#include <QMap>
#include <QStringList>
#include <QVector>

class QXmlStreamWriter;
class QXmlStreamReader;
class QIODevice;
class StylesTest;

namespace QXlsx {

class Format;

struct XlsxFormatNumberData
{
    XlsxFormatNumberData() : formatIndex(0) {}

    int formatIndex;
    QString formatString;
};

class XLSX_AUTOTEST_EXPORT Styles
{
public:
    Styles(bool createEmpty=false);
    ~Styles();
    void addXfFormat(const Format &format, bool force=false);
    Format xfFormat(int idx) const;
    void addDxfFormat(const Format &format, bool force=false);
    Format dxfFormat(int idx) const;

    QByteArray saveToXmlData();
    void saveToXmlFile(QIODevice *device);
    bool loadFromXmlFile(QIODevice *device);
    bool loadFromXmlData(const QByteArray &data);

private:
    friend class Format;
    friend class ::StylesTest;

    void fixNumFmt(const Format &format);

    void writeNumFmts(QXmlStreamWriter &writer);
    void writeFonts(QXmlStreamWriter &writer);
    void writeFont(QXmlStreamWriter &writer, const Format &font, bool isDxf = false);
    void writeFills(QXmlStreamWriter &writer);
    void writeFill(QXmlStreamWriter &writer, const Format &fill, bool isDxf = false);
    void writeBorders(QXmlStreamWriter &writer);
    void writeBorder(QXmlStreamWriter &writer, const Format &border, bool isDxf = false);
    void writeSubBorder(QXmlStreamWriter &writer, const QString &type, int style, const QColor &color, const QString &themeColor);
    void writeCellXfs(QXmlStreamWriter &writer);
    void writeDxfs(QXmlStreamWriter &writer);
    void writeDxf(QXmlStreamWriter &writer, const Format &format);

    bool readNumFmts(QXmlStreamReader &reader);
    bool readFonts(QXmlStreamReader &reader);
    bool readFont(QXmlStreamReader &reader, Format &format);
    bool readFills(QXmlStreamReader &reader);
    bool readFill(QXmlStreamReader &reader, Format &format);
    bool readBorders(QXmlStreamReader &reader);
    bool readBorder(QXmlStreamReader &reader, Format &format);
    bool readSubBorder(QXmlStreamReader &reader, const QString &name, Format::BorderStyle &style, QColor &color, QString &themeColor);
    bool readCellXfs(QXmlStreamReader &reader);
    bool readDxfs(QXmlStreamReader &reader);
    bool readDxf(QXmlStreamReader &reader);
    bool readColors(QXmlStreamReader &reader);
    bool readIndexedColors(QXmlStreamReader &reader);

    QColor getColorByIndex(int idx);

    QHash<QString, int> m_builtinNumFmtsHash;
    QMap<int, QSharedPointer<XlsxFormatNumberData> > m_customNumFmtIdMap;
    QHash<QString, QSharedPointer<XlsxFormatNumberData> > m_customNumFmtsHash;
    int m_nextCustomNumFmtId;
    QList<Format> m_fontsList;
    QList<Format> m_fillsList;
    QList<Format> m_bordersList;
    QHash<QByteArray, Format> m_fontsHash;
    QHash<QByteArray, Format> m_fillsHash;
    QHash<QByteArray, Format> m_bordersHash;

    QVector<QColor> m_indexedColors;

    QList<Format> m_xf_formatsList;
    QHash<QByteArray, Format> m_xf_formatsHash;

    QList<Format> m_dxf_formatsList;
    QHash<QByteArray, Format> m_dxf_formatsHash;

    bool m_emptyFormatAdded;
};

}
#endif // XLSXSTYLES_H
