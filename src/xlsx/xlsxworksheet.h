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
class CellRange;
struct XlsxImageData;
class RichString;

class WorksheetPrivate;
class Q_XLSX_EXPORT Worksheet
{
    Q_DECLARE_PRIVATE(Worksheet)
public:
    int write(const QString &row_column, const QVariant &value, const Format &format=Format());
    int write(int row, int column, const QVariant &value, const Format &format=Format());
    QVariant read(const QString &row_column) const;
    QVariant read(int row, int column) const;
    int writeString(const QString &row_column, const QString &value, const Format &format=Format());
    int writeString(int row, int column, const QString &value, const Format &format=Format());
    int writeString(const QString &row_column, const RichString &value, const Format &format=Format());
    int writeString(int row, int column, const RichString &value, const Format &format=Format());
    int writeInlineString(const QString &row_column, const QString &value, const Format &format=Format());
    int writeInlineString(int row, int column, const QString &value, const Format &format=Format());
    int writeNumeric(const QString &row_column, double value, const Format &format=Format());
    int writeNumeric(int row, int column, double value, const Format &format=Format());
    int writeFormula(const QString &row_column, const QString &formula, const Format &format=Format(), double result=0);
    int writeFormula(int row, int column, const QString &formula, const Format &format=Format(), double result=0);
    int writeArrayFormula(const QString &range, const QString &formula, const Format &format=Format());
    int writeArrayFormula(const CellRange &range, const QString &formula, const Format &format=Format());
    int writeBlank(const QString &row_column, const Format &format=Format());
    int writeBlank(int row, int column, const Format &format=Format());
    int writeBool(const QString &row_column, bool value, const Format &format=Format());
    int writeBool(int row, int column, bool value, const Format &format=Format());
    int writeDateTime(const QString &row_column, const QDateTime& dt, const Format &format=Format());
    int writeDateTime(int row, int column, const QDateTime& dt, const Format &format=Format());
    int writeHyperlink(const QString &row_column, const QUrl &url, const Format &format=Format(), const QString &display=QString(), const QString &tip=QString());
    int writeHyperlink(int row, int column, const QUrl &url, const Format &format=Format(), const QString &display=QString(), const QString &tip=QString());

    bool addDataValidation(const DataValidation &validation);

    Cell *cellAt(const QString &row_column) const;
    Cell *cellAt(int row, int column) const;

    int insertImage(int row, int column, const QImage &image, const QPointF &offset=QPointF(), double xScale=1, double yScale=1);

    int mergeCells(const QString &range, const Format &format=Format());
    int mergeCells(const CellRange &range, const Format &format=Format());
    int unmergeCells(const QString &range);
    int unmergeCells(const CellRange &range);

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
    void setHidden(bool hidden);
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
