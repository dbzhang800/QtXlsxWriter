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
#include "xlsxpackage_p.h"
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"
#include "xlsxcontenttypes_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxdocprops_p.h"
#include "xlsxtheme_p.h"
#include "xlsxstyles_p.h"
#include "xlsxrelationships_p.h"
#include "zipwriter_p.h"
#include <QBuffer>
#include <QDebug>

namespace QXlsx {

/*
    From Wikipedia: The Open Packaging Conventions (OPC) is a
    container-file technology initially created by Microsoft to store
    a combination of XML and non-XML files that together form a single
    entity such as an Open XML Paper Specification (OpenXPS)
    document. http://en.wikipedia.org/wiki/Open_Packaging_Conventions.

    At its simplest an Excel XLSX file contains the following elements:

         ____ [Content_Types].xml
        |
        |____ docProps
        | |____ app.xml
        | |____ core.xml
        |
        |____ xl
        | |____ workbook.xml
        | |____ worksheets
        | | |____ sheet1.xml
        | |
        | |____ styles.xml
        | |
        | |____ theme
        | | |____ theme1.xml
        | |
        | |_____rels
        | |____ workbook.xml.rels
        |
        |_____rels
          |____ .rels

    The Packager class coordinates the classes that represent the
    elements of the package and writes them into the XLSX file.
*/

Package::Package(Workbook *workbook) :
    m_workbook(workbook)
{
    m_worksheet_count = 0;
    m_chartsheet_count = 0;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->isChartsheet())
            m_chartsheet_count += 1;
        else
            m_worksheet_count += 1;
    }
}

bool Package::createPackage(const QString &packageName)
{
    QString fileName = packageName.isEmpty() ? m_workbook->fileName() : packageName;
    ZipWriter zipWriter(fileName);

    writeWorksheetFiles(zipWriter);
//    writeChartsheetFiles(zipWriter);
    writeWorkbookFile(zipWriter);
//    writeChartFiles(zipWriter);
//    writeDrawingFiles(zipWriter);
//    writeVmlFiles(zipWriter);
//    writeCommentFiles(zipWriter);
//    writeTableFiles(zipWriter);
    writeSharedStringsFile(zipWriter);
    writeDocPropsFiles(zipWriter);
    writeContentTypesFiles(zipWriter);
    writeStylesFiles(zipWriter);
    writeThemeFile(zipWriter);
    writeRootRelsFile(zipWriter);
    writeWorkbookRelsFile(zipWriter);
    writeWorksheetRelsFile(zipWriter);
//    writeChartsheetRelsFile(zipWriter);
//    writeImageFiles(zipWriter);
//    writeVbaProjectFiles(zipWriter);

    zipWriter.close();
    return true;
}

void Package::writeWorksheetFiles(ZipWriter &zipWriter)
{
    int index = 1;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->isChartsheet())
            continue;

        QByteArray data;
        QBuffer buffer(&data);
        buffer.open(QIODevice::WriteOnly);
        sheet->saveToXmlFile(&buffer);
        zipWriter.addFile(QString("xl/worksheets/sheet%1.xml").arg(index), data);
        index += 1;
    }
}

void Package::writeWorkbookFile(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    m_workbook->saveToXmlFile(&buffer);
    zipWriter.addFile("xl/workbook.xml", data);
}

void Package::writeContentTypesFiles(ZipWriter &zipWriter)
{
    ContentTypes content;

    int worksheet_index = 1;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->isChartsheet()) {

        } else {
            content.addWorksheetName(QString("sheet%1").arg(worksheet_index));
            worksheet_index += 1;
        }
    }

    if (m_workbook->sharedStrings()->count())
        content.addSharedString();

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    content.saveToXmlFile(&buffer);
    zipWriter.addFile("[Content_Types].xml", data);
}

void Package::writeDocPropsFiles(ZipWriter &zipWriter)
{
    DocProps props;

    if (m_worksheet_count)
        props.addHeadingPair("Worksheets", m_worksheet_count);
    if (m_chartsheet_count)
        props.addHeadingPair("Chartsheets", m_chartsheet_count);

    //Add worksheet parts
    foreach (Worksheet *sheet, m_workbook->worksheets()){
        if (!sheet->isChartsheet())
            props.addPartTitle(sheet->name());
    }

    //Add the chartsheet parts
    foreach (Worksheet *sheet, m_workbook->worksheets()){
        if (sheet->isChartsheet())
            props.addPartTitle(sheet->name());
    }

    QByteArray data1;
    QBuffer buffer1(&data1);
    buffer1.open(QIODevice::WriteOnly);
    props.saveToXmlFile_App(&buffer1);
    zipWriter.addFile("docProps/app.xml", data1);

    QByteArray data2;
    QBuffer buffer2(&data2);
    buffer2.open(QIODevice::WriteOnly);
    props.saveToXmlFile_Core(&buffer2);
    zipWriter.addFile("docProps/core.xml", data2);
}

void Package::writeSharedStringsFile(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    m_workbook->sharedStrings()->saveToXmlFile(&buffer);
    zipWriter.addFile("xl/sharedStrings.xml", data);
}

void Package::writeStylesFiles(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    m_workbook->styles()->saveToXmlFile(&buffer);
    zipWriter.addFile("xl/styles.xml", data);
}

void Package::writeThemeFile(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    Theme().saveToXmlFile(&buffer);
    zipWriter.addFile("xl/theme/theme1.xml", data);
}

void Package::writeRootRelsFile(ZipWriter &zipWriter)
{
    Relationships rels;
    rels.addDocumentRelationship("/officeDocument", "xl/workbook.xml");
    rels.addPackageRelationship("/metadata/core-properties", "docProps/core.xml");
    rels.addDocumentRelationship("/extended-properties", "docProps/app.xml");

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    rels.saveToXmlFile(&buffer);
    zipWriter.addFile("_rels/.rels", data);
}

void Package::writeWorkbookRelsFile(ZipWriter &zipWriter)
{
    Relationships rels;

    int worksheet_index = 1;
    int chartsheet_index = 1;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->isChartsheet()) {
            rels.addDocumentRelationship("/chartsheet", QString("chartsheets/sheet%1.xml").arg(chartsheet_index));
            chartsheet_index += 1;
        } else {
            rels.addDocumentRelationship("/worksheet", QString("worksheets/sheet%1.xml").arg(worksheet_index));
            worksheet_index += 1;
        }
    }

    rels.addDocumentRelationship("/theme", "theme/theme1.xml");
    rels.addDocumentRelationship("/styles", "styles.xml");

    if (m_workbook->sharedStrings()->count())
        rels.addDocumentRelationship("/sharedStrings", "sharedStrings.xml");

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    rels.saveToXmlFile(&buffer);
    zipWriter.addFile("xl/_rels/workbook.xml.rels", data);
}

void Package::writeWorksheetRelsFile(ZipWriter &zipWriter)
{

}
} // namespace QXlsx
