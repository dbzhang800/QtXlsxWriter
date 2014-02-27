/****************************************************************************
** Copyright (c) 2013-2014 Debao Zhang <hello@debao.me>
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

#include "xlsxabstractsheet.h"
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
class DocumentPrivate;
class Workbook;
class Format;
class Drawing;
class DataValidation;
class ConditionalFormatting;
class CellRange;
class RichString;
class Relationships;
class Chart;

class WorksheetPrivate;
class Q_XLSX_EXPORT Worksheet : public AbstractSheet
{
    Q_DECLARE_PRIVATE(Worksheet)
public:
    bool write(const QString &row_column, const QVariant &value, const Format &format=Format());
    bool write(int row, int column, const QVariant &value, const Format &format=Format());
    QVariant read(const QString &row_column) const;
    QVariant read(int row, int column) const;
    bool writeString(const QString &row_column, const QString &value, const Format &format=Format());
    bool writeString(int row, int column, const QString &value, const Format &format=Format());
    bool writeString(const QString &row_column, const RichString &value, const Format &format=Format());
    bool writeString(int row, int column, const RichString &value, const Format &format=Format());
    bool writeInlineString(const QString &row_column, const QString &value, const Format &format=Format());
    bool writeInlineString(int row, int column, const QString &value, const Format &format=Format());
    bool writeNumeric(const QString &row_column, double value, const Format &format=Format());
    bool writeNumeric(int row, int column, double value, const Format &format=Format());
    bool writeFormula(const QString &row_column, const QString &formula, const Format &format=Format(), double result=0);
    bool writeFormula(int row, int column, const QString &formula, const Format &format=Format(), double result=0);
    bool writeArrayFormula(const QString &range, const QString &formula, const Format &format=Format());
    bool writeArrayFormula(const CellRange &range, const QString &formula, const Format &format=Format());
    bool writeBlank(const QString &row_column, const Format &format=Format());
    bool writeBlank(int row, int column, const Format &format=Format());
    bool writeBool(const QString &row_column, bool value, const Format &format=Format());
    bool writeBool(int row, int column, bool value, const Format &format=Format());
    bool writeDateTime(const QString &row_column, const QDateTime& dt, const Format &format=Format());
    bool writeDateTime(int row, int column, const QDateTime& dt, const Format &format=Format());
    bool writeTime(const QString &row_column, const QTime& t, const Format &format=Format());
    bool writeTime(int row, int column, const QTime& t, const Format &format=Format());

    bool writeHyperlink(const QString &row_column, const QUrl &url, const Format &format=Format(), const QString &display=QString(), const QString &tip=QString());
    bool writeHyperlink(int row, int column, const QUrl &url, const Format &format=Format(), const QString &display=QString(), const QString &tip=QString());

    bool addDataValidation(const DataValidation &validation);
    bool addConditionalFormatting(const ConditionalFormatting &cf);

    Cell *cellAt(const QString &row_column) const;
    Cell *cellAt(int row, int column) const;

    bool insertImage(int row, int column, const QImage &image);
    Chart *insertChart(int row, int column, const QSize &size);

    bool mergeCells(const QString &range, const Format &format=Format());
    bool mergeCells(const CellRange &range, const Format &format=Format());
    bool unmergeCells(const QString &range);
    bool unmergeCells(const CellRange &range);
    QList<CellRange> mergedCells() const;

    bool setRow(int row, double height, const Format &format=Format(), bool hidden=false);
    bool setColumn(int colFirst, int colLast, double width, const Format &format=Format(), bool hidden=false);
    bool setColumn(const QString &colFirst, const QString &colLast, double width, const Format &format=Format(), bool hidden=false);
    bool groupRows(int rowFirst, int rowLast, bool collapsed = true);
    bool groupColumns(int colFirst, int colLast, bool collapsed = true);
    bool groupColumns(const QString &colFirst, const QString &colLast, bool collapsed = true);
    CellRange dimension() const;

    bool isWindowProtected() const;
    void setWindowProtected(bool protect);
    bool isFormulasVisible() const;
    void setFormulasVisible(bool visible);
    bool isGridLinesVisible() const;
    void setGridLinesVisible(bool visible);
    bool isRowColumnHeadersVisible() const;
    void setRowColumnHeadersVisible(bool visible);
    bool isZerosVisible() const;
    void setZerosVisible(bool visible);
    bool isRightToLeft() const;
    void setRightToLeft(bool enable);
    bool isSelected() const;
    void setSelected(bool select);
    bool isRulerVisible() const;
    void setRulerVisible(bool visible);
    bool isOutlineSymbolsVisible() const;
    void setOutlineSymbolsVisible(bool visible);
    bool isWhiteSpaceVisible() const;
    void setWhiteSpaceVisible(bool visible);

    ~Worksheet();
private:
    friend class DocumentPrivate;
    friend class Workbook;
    friend class ::WorksheetTest;
    Worksheet(const QString &sheetName, int sheetId, Workbook *book, CreateFlag flag);
    Worksheet *copy(const QString &distName, int distId) const;

    void saveToXmlFile(QIODevice *device) const;
    bool loadFromXmlFile(QIODevice *device);
};

QT_END_NAMESPACE_XLSX
#endif // XLSXWORKSHEET_H
