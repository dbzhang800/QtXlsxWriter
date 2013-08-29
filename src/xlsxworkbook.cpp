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
#include "xmlstreamwriter_p.h"

namespace QXlsx {

WorkbookPrivate::WorkbookPrivate(Workbook *q) :
    q_ptr(q)
{
    sharedStrings = new SharedStrings(q);
    styles = new Styles(q);

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

Workbook::Workbook(QObject *parent) :
    QObject(parent), d_ptr(new WorkbookPrivate(this))
{

}

Workbook::~Workbook()
{
    delete d_ptr;
}

void Workbook::save(const QString &name)
{
    Q_D(Workbook);

    //Add a default worksheet if non have been added.
    if (d->worksheets.size() == 0)
        addWorksheet();

    //Ensure that at least one worksheet has been selected.
    if (d->activesheet == 0) {
        d->worksheets[0]->setHidden(false);
        d->worksheets[0]->setSelected(true);
    }

    //Set the active sheet
    foreach (Worksheet *sheet, d->worksheets) {
        if (sheet->index() == d->activesheet)
            sheet->setActived(true);
    }

    //Create the package based on current workbook
    Package package(this);
    package.createPackage(name);
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

    QString worksheetName = name;
    int index = d->worksheets.size()+1;
    if (name.isEmpty())
        worksheetName = QString("Sheet%1").arg(index);

    Worksheet *sheet = new Worksheet(worksheetName, index, this);
    d->worksheets.append(sheet);
    return sheet;
}

Format *Workbook::addFormat()
{
    Q_D(Workbook);
    return d->styles->addFormat();
}

QList<Worksheet *> Workbook::worksheets() const
{
    Q_D(const Workbook);
    return d->worksheets;
}

SharedStrings *Workbook::sharedStrings()
{
    Q_D(Workbook);
    return d->sharedStrings;
}

Styles *Workbook::styles()
{
    Q_D(Workbook);
    return d->styles;
}

void Workbook::saveToXmlFile(QIODevice *device)
{
    Q_D(Workbook);
    XmlStreamWriter writer(device);

    writer.writeStartDocument("1.0", true);
    writer.writeStartElement("workbook");
    writer.writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer.writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

    writer.writeEmptyElement("fileVersion");
    writer.writeAttribute("appName", "xl");
    writer.writeAttribute("lastEdited", "4");
    writer.writeAttribute("lowestEdited", "4");
    writer.writeAttribute("rupBuild", "4505");
//    writer.writeAttribute("codeName", "{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}");

    writer.writeEmptyElement("workbookPr");
    if (d->date1904)
        writer.writeAttribute("date1904", "1");
    writer.writeAttribute("defaultThemeVersion", "124226");

    writer.writeStartElement("bookViews");
    writer.writeEmptyElement("workbookView");
    writer.writeAttribute("xWindow", QString::number(d->x_window));
    writer.writeAttribute("yWindow", QString::number(d->y_window));
    writer.writeAttribute("windowWidth", QString::number(d->window_width));
    writer.writeAttribute("windowHeight", QString::number(d->window_height));
    //Store the firstSheet when it isn't the default
    if (d->firstsheet > 0)
        writer.writeAttribute("firstSheet", QString::number(d->firstsheet + 1));
    //Store the activeTab when it isn't the first sheet
    if (d->activesheet > 0)
        writer.writeAttribute("activeTab", QString::number(d->activesheet));
    writer.writeEndElement();//bookviews

    writer.writeStartElement("sheets");
    foreach (Worksheet *sheet, d->worksheets) {
        writer.writeEmptyElement("sheet");
        writer.writeAttribute("name", sheet->name());
        writer.writeAttribute("sheetId", QString::number(sheet->index()));
        if (sheet->isHidden())
            writer.writeAttribute("state", "hidden");
        writer.writeAttribute("r:id", QString("rId%1").arg(sheet->index()));
    }
    writer.writeEndElement();//sheets

//    writer.writeStartElement("definedNames");
//    writer.writeEndElement();//definedNames

    writer.writeStartElement("calcPr");
    writer.writeAttribute("calcId", "124519");
    writer.writeEndElement(); //calcPr

    writer.writeEndElement();//workbook
    writer.writeEndDocument();
}

} //namespace
