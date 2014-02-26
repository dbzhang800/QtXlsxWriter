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

#ifndef QXLSX_XLSXDOCUMENT_H
#define QXLSX_XLSXDOCUMENT_H

#include "xlsxglobal.h"
#include "xlsxformat.h"
#include "xlsxworksheet.h"
#include <QObject>
#include <QVariant>
class QIODevice;
class QImage;

QT_BEGIN_NAMESPACE_XLSX

class Workbook;
class Cell;
class CellRange;
class DataValidation;
class ConditionalFormatting;
class Chart;

class DocumentPrivate;
class Q_XLSX_EXPORT Document : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Document)

public:
    explicit Document(QObject *parent = 0);
    Document(const QString &xlsxName, QObject *parent=0);
    Document(QIODevice *device, QObject *parent=0);
    ~Document();

    bool write(const QString &cell, const QVariant &value, const Format &format=Format());
    bool write(int row, int col, const QVariant &value, const Format &format=Format());
    QVariant read(const QString &cell) const;
    QVariant read(int row, int col) const;
    bool insertImage(int row, int col, const QImage &image);
    Chart *insertChart(int row, int col, const QSize &size);
    bool mergeCells(const CellRange &range, const Format &format=Format());
    bool mergeCells(const QString &range, const Format &format=Format());
    bool unmergeCells(const CellRange &range);
    bool unmergeCells(const QString &range);
    bool setRow(int row, double height, const Format &format=Format(), bool hidden=false);
    bool setColumn(int colFirst, int colLast, double width, const Format &format=Format(), bool hidden=false);
    bool setColumn(const QString &colFirst, const QString &colLast, double width, const Format &format=Format(), bool hidden=false);
    bool groupRows(int rowFirst, int rowLast, bool collapsed = true);
    bool groupColumns(int colFirst, int colLast, bool collapsed = true);
    bool addDataValidation(const DataValidation &validation);
    bool addConditionalFormatting(const ConditionalFormatting &cf);

    Cell *cellAt(const QString &cell) const;
    Cell *cellAt(int row, int col) const;

    bool defineName(const QString &name, const QString &formula, const QString &comment=QString(), const QString &scope=QString());

    CellRange dimension() const;

    QString documentProperty(const QString &name) const;
    void setDocumentProperty(const QString &name, const QString &property);
    QStringList documentPropertyNames() const;

    QStringList sheetNames() const;
    bool addSheet(const QString &name = QString(), AbstractSheet::SheetType type = AbstractSheet::ST_WorkSheet);
    bool insertSheet(int index, const QString &name = QString(), AbstractSheet::SheetType type = AbstractSheet::ST_WorkSheet);
    bool selectSheet(const QString &name);
    bool renameSheet(const QString &oldName, const QString &newName);
    bool copySheet(const QString &srcName, const QString &distName = QString());
    bool moveSheet(const QString &srcName, int distIndex);
    bool deleteSheet(const QString &name);

    Workbook *workbook() const;
    AbstractSheet *sheet(const QString &sheetName) const;
    AbstractSheet *currentSheet() const;
    Worksheet *currentWorksheet() const;

    bool save() const;
    bool saveAs(const QString &xlsXname) const;
    bool saveAs(QIODevice *device) const;

private:
    Q_DISABLE_COPY(Document)
    DocumentPrivate * const d_ptr;
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_XLSXDOCUMENT_H
