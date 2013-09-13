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
#include "xlsxworkbook.h"
#include "xlsxworkbook_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxworksheet.h"
#include "xlsxstyles_p.h"
#include "xlsxformat.h"
#include "xlsxpackage_p.h"
#include "xlsxxmlwriter_p.h"
#include "xlsxxmlreader_p.h"
#include "xlsxworksheet_p.h"

#include <QFile>
#include <QBuffer>

namespace QXlsx {

WorkbookPrivate::WorkbookPrivate(Workbook *q) :
    q_ptr(q)
{
    sharedStrings = QSharedPointer<SharedStrings> (new SharedStrings);
    styles = QSharedPointer<Styles>(new Styles);

    x_window = 240;
    y_window = 15;
    window_width = 16095;
    window_height = 9660;

    strings_to_numbers_enabled = false;
    date1904 = false;
    activesheet = 0;
    firstsheet = 0;
    table_count = 0;
}

Workbook::Workbook() :
    d_ptr(new WorkbookPrivate(this))
{

}

Workbook::~Workbook()
{
    delete d_ptr;
}

bool Workbook::isDate1904() const
{
    Q_D(const Workbook);
    return d->date1904;
}

/*
 Excel for Windows uses a default epoch of 1900 and Excel
 for Mac uses an epoch of 1904. However, Excel on either
 platform will convert automatically between one system
 and the other. QtXlsxWriter stores dates in the 1900 format
 by default.
*/
void Workbook::setDate1904(bool date1904)
{
    Q_D(Workbook);
    d->date1904 = date1904;
}

/*
  Enable the worksheet.write() method to convert strings
  to numbers, where possible, using float() in order to avoid
  an Excel warning about "Numbers Stored as Text".

  The default is false
 */
void Workbook::setStringsToNumbersEnabled(bool enable)
{
    Q_D(Workbook);
    d->strings_to_numbers_enabled = enable;
}

bool Workbook::isStringsToNumbersEnabled() const
{
    Q_D(const Workbook);
    return d->strings_to_numbers_enabled;
}

void Workbook::defineName(const QString &name, const QString &formula)
{

}

Worksheet *Workbook::addWorksheet(const QString &name)
{
    Q_D(Workbook);
    return insertWorkSheet(d->worksheets.size(), name);
}

Worksheet *Workbook::insertWorkSheet(int index, const QString &name)
{
    Q_D(Workbook);
    static int lastIndex = -1;
    QString worksheetName = name;
    if (!name.isEmpty()) {
        //If user given an already in-use name, we should not continue any more!
        for (int i=0; i<d->worksheets.size(); ++i) {
            if (d->worksheets[i]->sheetName() == name) {
                return 0;
            }
        }
    } else {
        bool exists;
        do {
            ++lastIndex;
            exists = false;
            worksheetName = QStringLiteral("Sheet%1").arg(lastIndex+1);
            for (int i=0; i<d->worksheets.size(); ++i) {
                if (d->worksheets[i]->sheetName() == worksheetName)
                    exists = true;
            }
        } while (exists);
    }

    Worksheet *sheet = new Worksheet(worksheetName, this);
    d->worksheets.insert(index, QSharedPointer<Worksheet>(sheet));
    d->activesheet = index;
    return sheet;
}

int Workbook::activedWorksheet() const
{
    Q_D(const Workbook);
    return d->activesheet;
}

void Workbook::setActivedWorksheet(int index)
{
    Q_D(Workbook);
    if (index < 0 || index >= d->worksheets.size()) {
        //warning
        return;
    }
    d->activesheet = index;
}

Format *Workbook::createFormat()
{
    Q_D(Workbook);
    return d->styles->createFormat();
}

QList<QSharedPointer<Worksheet> > Workbook::worksheets() const
{
    Q_D(const Workbook);
    return d->worksheets;
}

SharedStrings *Workbook::sharedStrings()
{
    Q_D(Workbook);
    return d->sharedStrings.data();
}

Styles *Workbook::styles()
{
    Q_D(Workbook);
    return d->styles.data();
}

QList<QImage> Workbook::images()
{
    Q_D(Workbook);
    return d->images;
}

QList<Drawing *> Workbook::drawings()
{
    Q_D(Workbook);
    return d->drawings;
}

void Workbook::prepareDrawings()
{
    Q_D(Workbook);
    int drawing_id = 0;
    int image_ref_id = 0;
    d->images.clear();
    d->drawings.clear();

    for (int i=0; i<d->worksheets.size(); ++i) {
        QSharedPointer<Worksheet> sheet = d->worksheets[i];
        if (sheet->images().isEmpty()) //No drawing (such as Image, ...)
            continue;

        sheet->clearExtraDrawingInfo();

        //At present, only picture type supported
        drawing_id += 1;
        for (int idx = 0; idx < sheet->images().size(); ++idx) {
            image_ref_id += 1;
            sheet->prepareImage(idx, image_ref_id, drawing_id);
            d->images.append(sheet->images()[idx]->image);
        }

        d->drawings.append(sheet->drawing());
    }
}

void Workbook::saveToXmlFile(QIODevice *device)
{
    Q_D(Workbook);
    XmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("workbook"));
    writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    writer.writeAttribute(QStringLiteral("xmlns:r"), QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships"));

    writer.writeEmptyElement(QStringLiteral("fileVersion"));
    writer.writeAttribute(QStringLiteral("appName"), QStringLiteral("xl"));
    writer.writeAttribute(QStringLiteral("lastEdited"), QStringLiteral("4"));
    writer.writeAttribute(QStringLiteral("lowestEdited"), QStringLiteral("4"));
    writer.writeAttribute(QStringLiteral("rupBuild"), QStringLiteral("4505"));
//    writer.writeAttribute(QStringLiteral("codeName"), QStringLiteral("{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}"));

    writer.writeEmptyElement(QStringLiteral("workbookPr"));
    if (d->date1904)
        writer.writeAttribute(QStringLiteral("date1904"), QStringLiteral("1"));
    writer.writeAttribute(QStringLiteral("defaultThemeVersion"), QStringLiteral("124226"));

    writer.writeStartElement(QStringLiteral("bookViews"));
    writer.writeEmptyElement(QStringLiteral("workbookView"));
    writer.writeAttribute(QStringLiteral("xWindow"), QString::number(d->x_window));
    writer.writeAttribute(QStringLiteral("yWindow"), QString::number(d->y_window));
    writer.writeAttribute(QStringLiteral("windowWidth"), QString::number(d->window_width));
    writer.writeAttribute(QStringLiteral("windowHeight"), QString::number(d->window_height));
    //Store the firstSheet when it isn't the default
    if (d->firstsheet > 0)
        writer.writeAttribute(QStringLiteral("firstSheet"), QString::number(d->firstsheet + 1));
    //Store the activeTab when it isn't the first sheet
    if (d->activesheet > 0)
        writer.writeAttribute(QStringLiteral("activeTab"), QString::number(d->activesheet));
    writer.writeEndElement();//bookviews

    writer.writeStartElement(QStringLiteral("sheets"));
    for (int i=0; i<d->worksheets.size(); ++i) {
        QSharedPointer<Worksheet> sheet = d->worksheets[i];
        writer.writeEmptyElement(QStringLiteral("sheet"));
        writer.writeAttribute(QStringLiteral("name"), sheet->sheetName());
        writer.writeAttribute(QStringLiteral("sheetId"), QString::number(i+1));
        if (sheet->isHidden())
            writer.writeAttribute(QStringLiteral("state"), QStringLiteral("hidden"));
        writer.writeAttribute(QStringLiteral("r:id"), QStringLiteral("rId%1").arg(i+1));
    }
    writer.writeEndElement();//sheets

//    writer.writeStartElement(QStringLiteral("definedNames"));
//    writer.writeEndElement();//definedNames

    writer.writeStartElement(QStringLiteral("calcPr"));
    writer.writeAttribute(QStringLiteral("calcId"), QStringLiteral("124519"));
    writer.writeEndElement(); //calcPr

    writer.writeEndElement();//workbook
    writer.writeEndDocument();
}

QByteArray Workbook::saveToXmlData()
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    saveToXmlFile(&buffer);

    return data;
}

QSharedPointer<Workbook> Workbook::loadFromXmlFile(QIODevice *device)
{
    Workbook *book = new Workbook;

    XmlStreamReader reader(device);
    while(!reader.atEnd()) {
         QXmlStreamReader::TokenType token = reader.readNext();
         if (token == QXmlStreamReader::StartElement) {
             if (reader.name() == QLatin1String("sheet")) {
                 QXmlStreamAttributes attributes = reader.attributes();
                 QString sheetName = attributes.value(QLatin1String("name")).toString();
                 QString rId = attributes.value(QLatin1String("r:id")).toString();
                 book->d_func()->sheetNameIdPairList.append(QPair<QString, QString>(sheetName, rId));
             }
         }
    }
    return QSharedPointer<Workbook>(book);
}

QSharedPointer<Workbook> Workbook::loadFromXmlData(const QByteArray &data)
{
    QBuffer buffer;
    buffer.setData(data);
    buffer.open(QIODevice::ReadOnly);

    return loadFromXmlFile(&buffer);
}

void Workbook::addWorksheet(const QString &name, QSharedPointer<Worksheet> sheet)
{
    Q_D(Workbook);

    sheet->setSheetName(name);
    d->worksheets.append(sheet);
}

} //namespace
