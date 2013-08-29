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
#ifndef XLSXWORKBOOK_H
#define XLSXWORKBOOK_H

#include <QObject>
#include <QList>
class QIODevice;

namespace QXlsx {

class Worksheet;
class Format;
class SharedStrings;
class Styles;
class Package;

class Workbook : public QObject
{
    Q_OBJECT
public:
    Workbook(QObject *parent=0);
    ~Workbook();

    Worksheet *addWorksheet(const QString &name = QString());
    Format *addFormat();
//    void addChart();
    void defineName(const QString &name, const QString &formula);
    bool isDate1904() const;
    void setDate1904(bool date1904);
    bool isStringsToNumbersEnabled() const;
    void setStringsToNumbersEnabled(bool enable=true);

    void save(const QString &name);

private:
    friend class Package;
    friend class Worksheet;

    QList<Worksheet *> worksheets() const;
    SharedStrings *sharedStrings();
    Styles *styles();
    void saveToXmlFile(QIODevice *device);

    SharedStrings *m_sharedStrings;
    QList<Worksheet *> m_worksheets;
    Styles *m_styles;
    bool m_strings_to_numbers_enabled;
    bool m_date1904;

    int m_x_window;
    int m_y_window;
    int m_window_width;
    int m_window_height;

    int m_activesheet;
    int m_firstsheet;
    int m_table_count;
};

} //QXlsx

#endif // XLSXWORKBOOK_H
