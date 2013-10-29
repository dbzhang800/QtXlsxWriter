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

#include "xlsxglobal.h"
#include "xlsxcell.h"
#include "xlsxcellrange.h"
#include <QStringList>
#include <QMap>
#include <QVariant>
#include <QPointF>
#include <QSharedPointer>
class QIODevice;
class QDateTime;
class QUrl;
class QImage;
class WorksheetTest;

QT_BEGIN_NAMESPACE_XLSX
class Package;
class Workbook;
class Format;
class Drawing;
class DataValidation;
struct XlsxImageData;

class WorksheetPrivate;
class Q_XLSX_EXPORT Worksheet
{
    Q_DECLARE_PRIVATE(Worksheet)
public:
    int write(const QString &row_column, const QVariant &value, Format *format=0);
    int write(int row, int column, const QVariant &value, Format *format=0);
    int writeString(int row, int column, const QString &value, Format *format=0);
    int writeInlineString(int row, int column, const QString &value, Format *format=0);
    int writeNumeric(int row, int column, double value, Format *format=0);
    int writeFormula(int row, int column, const QString &formula, Format *format=0, double result=0);
    int writeBlank(int row, int column, Format *format=0);
    int writeBool(int row, int column, bool value, Format *format=0);
    int writeDateTime(int row, int column, const QDateTime& dt, Format *format=0);
    int writeHyperlink(int row, int column, const QUrl &url, Format *format=0, const QString &display=QString(), const QString &tip=QString());

    bool addDataValidation(const DataValidation &validation);

    Cell *cellAt(const QString &row_column) const;
    Cell *cellAt(int row, int column) const;

    int insertImage(int row, int column, const QImage &image, const QPointF &offset=QPointF(), double xScale=1, double yScale=1);

    int mergeCells(int row_begin, int column_begin, int row_end, int column_end);
    int mergeCells(const QString &range);
    int mergeCells(const CellRange &range);
    int unmergeCells(int row_begin, int column_begin, int row_end, int column_end);
    int unmergeCells(const QString &range);
    int unmergeCells(const CellRange &range);

    bool setRow(int row, double height, Format* format=0, bool hidden=false);
    bool setColumn(int colFirst, int colLast, double width, Format* format=0, bool hidden=false);

    int firstRow() const;
    int lastRow() const;
    int firstColumn() const;
    int lastColumn() const;

    void setRightToLeft(bool enable);
    void setZeroValuesHidden(bool enable);

    QString sheetName() const;
    void setSheetName(const QString &sheetName);

    Workbook *workbook() const;
    ~Worksheet();
private:
    friend class Package;
    friend class Workbook;
    friend class ::WorksheetTest;
    Worksheet(const QString &sheetName, int sheetId, Workbook *book);

    void saveToXmlFile(QIODevice *device);
    QByteArray saveToXmlData();
    bool loadFromXmlFile(QIODevice *device);
    bool loadFromXmlData(const QByteArray &data);

    bool isChartsheet() const;
    bool isHidden() const;
    bool isSelected() const;
    void setHidden(bool hidden);
    void setSelected(bool select);
    int sheetId() const;
    QStringList externUrlList() const;
    QStringList externDrawingList() const;
    QList<QPair<QString, QString> > drawingLinks() const;
    Drawing *drawing() const;
    QList<XlsxImageData *> images() const;
    void prepareImage(int index, int image_id, int drawing_id);
    void clearExtraDrawingInfo();

    WorksheetPrivate * const d_ptr;
};

QT_END_NAMESPACE_XLSX
#endif // XLSXWORKSHEET_H
