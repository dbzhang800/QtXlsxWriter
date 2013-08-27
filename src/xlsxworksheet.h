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
#ifndef XLSXWORKSHEET_H
#define XLSXWORKSHEET_H

#include <QObject>
#include <QStringList>
#include <QMap>
#include <QVariant>
class QIODevice;

namespace QXlsx {
class Package;
class Workbook;
class XmlStreamWriter;
class Format;

struct XlsxCellData
{
    enum CellDataType {
        Blank,
        String,
        Number,
        Formula,
        ArrayFormula,
        Boolean
    };
    XlsxCellData(const QVariant &data=QVariant(), CellDataType type=Blank, Format *format=0) :
        value(data), dataType(type), format(format)
    {

    }

    QVariant value;
    QString formula;
    CellDataType dataType;
    Format *format;
};

class Worksheet : public QObject
{
    Q_OBJECT
public:
    int write(const QString row_column, const QVariant &value, Format *format=0);
    int write(int row, int column, const QVariant &value, Format *format=0);
    int writeString(int row, int column, const QString &value, Format *format=0);
    int writeNumber(int row, int column, double value, Format *format=0);
    int writeFormula(int row, int column, const QString &formula, Format *format=0, double result=0);
    int writeBlank(int row, int column, Format *format=0);
    int writeBool(int row, int column, bool value, Format *format=0);

    void setRightToLeft(bool enable);
    void setZeroValuesHidden(bool enable);

private:
    friend class Package;
    friend class Workbook;
    explicit Worksheet(const QString &sheetName, int sheetIndex, Workbook *parent=0);
    virtual bool isChartsheet() const;
    QString name() const;
    int index() const;
    bool isHidden() const;
    bool isSelected() const;
    bool isActived() const;
    void setHidden(bool hidden);
    void setSelected(bool select);
    void setActived(bool act);
    void saveToXmlFile(QIODevice *device);
    int checkDimensions(int row, int col, bool ignore_row=false, bool ignore_col=false);
    QString generateDimensionString();
    void writeSheetData(XmlStreamWriter &writer);
    void writeCellData(XmlStreamWriter &writer, int row, int col, const XlsxCellData &data);
    void calculateSpans();

    Workbook *m_workbook;
    QMap<int, QMap<int, XlsxCellData> > m_table;
    QMap<int, QMap<int, QString> > m_comments;

    int m_xls_rowmax;
    int m_xls_colmax;
    int m_xls_strmax;
    int m_dim_rowmin;
    int m_dim_rowmax;
    int m_dim_colmin;
    int m_dim_colmax;
    int m_previous_row;

    QMap<int, QString> m_row_spans;

    int m_outline_row_level;
    int m_outline_col_level;

    int m_default_row_height;
    bool m_default_row_zeroed;

    QString m_name;
    int m_index;
    bool m_hidden;
    bool m_selected;
    bool m_actived;
    bool m_right_to_left;
    bool m_show_zeros;
};

} //QXlsx

Q_DECLARE_TYPEINFO(QXlsx::XlsxCellData, Q_MOVABLE_TYPE);

#endif // XLSXWORKSHEET_H
