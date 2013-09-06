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
#include "xlsxdocpropscore_p.h"
#include "xlsxdocpropsapp_p.h"
#include "xlsxtheme_p.h"
#include "xlsxstyles_p.h"
#include "xlsxrelationships_p.h"
#include "xlsxzipwriter_p.h"
#include "xlsxdrawing_p.h"
#include "xlsxzipreader_p.h"
#include "xlsxdocument.h"
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

bool Package::parsePackage(QIODevice *packageDevice, Document *document)
{
    ZipReader zipReader(packageDevice);
    QStringList filePaths = zipReader.filePaths();

    return false;
}

bool Package::createPackage(const QString &packageName)
{
    ZipWriter zipWriter(packageName);
    if (zipWriter.error())
        return false;

    m_workbook->styles()->clearExtraFormatInfo(); //These info will be generated when write the worksheet data.
    m_workbook->prepareDrawings();

    writeWorksheetFiles(zipWriter);
//    writeChartsheetFiles(zipWriter);
    writeWorkbookFile(zipWriter);
//    writeChartFiles(zipWriter);
    writeDrawingFiles(zipWriter);
//    writeVmlFiles(zipWriter);
//    writeCommentFiles(zipWriter);
//    writeTableFiles(zipWriter);
    writeSharedStringsFile(zipWriter);
    writeDocPropsAppFile(zipWriter);
    writeDocPropsCoreFile(zipWriter);
    writeContentTypesFile(zipWriter);
    m_workbook->styles()->prepareStyles();
    writeStylesFiles(zipWriter);
    writeThemeFile(zipWriter);
    writeRootRelsFile(zipWriter);
    writeWorkbookRelsFile(zipWriter);
    writeWorksheetRelsFiles(zipWriter);
//    writeChartsheetRelsFile(zipWriter);
    writeDrawingRelsFiles(zipWriter);
    writeImageFiles(zipWriter);
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
        zipWriter.addFile(QStringLiteral("xl/worksheets/sheet%1.xml").arg(index), data);
        index += 1;
    }
}

void Package::writeWorkbookFile(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    m_workbook->saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("xl/workbook.xml"), data);
}

void Package::writeDrawingFiles(ZipWriter &zipWriter)
{
    for (int i=0; i<m_workbook->drawings().size(); ++i) {
        Drawing *drawing = m_workbook->drawings()[i];

        QByteArray data;
        QBuffer buffer(&data);
        buffer.open(QIODevice::WriteOnly);
        drawing->saveToXmlFile(&buffer);
        zipWriter.addFile(QStringLiteral("xl/drawings/drawing%1.xml").arg(i+1), data);
    }
}

void Package::writeContentTypesFile(ZipWriter &zipWriter)
{
    ContentTypes content;

    int worksheet_index = 1;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->isChartsheet()) {

        } else {
            content.addWorksheetName(QStringLiteral("sheet%1").arg(worksheet_index));
            worksheet_index += 1;
        }
    }

    int drawing_index = 1;
    foreach (Drawing *drawing, m_workbook->drawings()) {
        content.addDrawingName(QStringLiteral("drawing%1").arg(drawing_index));
        drawing_index += 1;
    }

    if (!m_workbook->images().isEmpty())
        content.addImageTypes(QStringList()<<QStringLiteral("png"));

    if (m_workbook->sharedStrings()->count())
        content.addSharedString();

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    content.saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("[Content_Types].xml"), data);
}

void Package::writeDocPropsAppFile(ZipWriter &zipWriter)
{
    DocPropsApp props;

    foreach (QByteArray name, m_workbook->dynamicPropertyNames())
        props.setProperty(name.data(), m_workbook->property(name.data()));

    if (m_worksheet_count)
        props.addHeadingPair(QStringLiteral("Worksheets"), m_worksheet_count);
    if (m_chartsheet_count)
        props.addHeadingPair(QStringLiteral("Chartsheets"), m_chartsheet_count);

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

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    props.saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("docProps/app.xml"), data);
}

void Package::writeDocPropsCoreFile(ZipWriter &zipWriter)
{
    DocPropsCore props;

    foreach (QByteArray name, m_workbook->dynamicPropertyNames())
        props.setProperty(name.data(), m_workbook->property(name.data()));

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    props.saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("docProps/core.xml"), data);
}

void Package::writeSharedStringsFile(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    m_workbook->sharedStrings()->saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("xl/sharedStrings.xml"), data);
}

void Package::writeStylesFiles(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    m_workbook->styles()->saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("xl/styles.xml"), data);
}

void Package::writeThemeFile(ZipWriter &zipWriter)
{
    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    Theme().saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("xl/theme/theme1.xml"), data);
}

void Package::writeRootRelsFile(ZipWriter &zipWriter)
{
    Relationships rels;
    rels.addDocumentRelationship(QStringLiteral("/officeDocument"), QStringLiteral("xl/workbook.xml"));
    rels.addPackageRelationship(QStringLiteral("/metadata/core-properties"), QStringLiteral("docProps/core.xml"));
    rels.addDocumentRelationship(QStringLiteral("/extended-properties"), QStringLiteral("docProps/app.xml"));

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    rels.saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("_rels/.rels"), data);
}

void Package::writeWorkbookRelsFile(ZipWriter &zipWriter)
{
    Relationships rels;

    int worksheet_index = 1;
    int chartsheet_index = 1;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->isChartsheet()) {
            rels.addDocumentRelationship(QStringLiteral("/chartsheet"), QStringLiteral("chartsheets/sheet%1.xml").arg(chartsheet_index));
            chartsheet_index += 1;
        } else {
            rels.addDocumentRelationship(QStringLiteral("/worksheet"), QStringLiteral("worksheets/sheet%1.xml").arg(worksheet_index));
            worksheet_index += 1;
        }
    }

    rels.addDocumentRelationship(QStringLiteral("/theme"), QStringLiteral("theme/theme1.xml"));
    rels.addDocumentRelationship(QStringLiteral("/styles"), QStringLiteral("styles.xml"));

    if (m_workbook->sharedStrings()->count())
        rels.addDocumentRelationship(QStringLiteral("/sharedStrings"), QStringLiteral("sharedStrings.xml"));

    QByteArray data;
    QBuffer buffer(&data);
    buffer.open(QIODevice::WriteOnly);
    rels.saveToXmlFile(&buffer);
    zipWriter.addFile(QStringLiteral("xl/_rels/workbook.xml.rels"), data);
}

void Package::writeWorksheetRelsFiles(ZipWriter &zipWriter)
{
    int index = 1;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->isChartsheet())
            continue;
        Relationships rels;

        foreach (QString link, sheet->externUrlList())
            rels.addWorksheetRelationship(QStringLiteral("/hyperlink"), link, QStringLiteral("External"));
        foreach (QString link, sheet->externDrawingList())
            rels.addWorksheetRelationship(QStringLiteral("/drawing"), link);
        QByteArray data;
        QBuffer buffer(&data);
        buffer.open(QIODevice::WriteOnly);
        rels.saveToXmlFile(&buffer);
        zipWriter.addFile(QStringLiteral("xl/worksheets/_rels/sheet%1.xml.rels").arg(index), data);
        index += 1;
    }
}

void Package::writeDrawingRelsFiles(ZipWriter &zipWriter)
{
    int index = 1;
    foreach (Worksheet *sheet, m_workbook->worksheets()) {
        if (sheet->drawingLinks().size() == 0)
            continue;
        Relationships rels;

        typedef QPair<QString, QString> PairType;
        foreach (PairType pair, sheet->drawingLinks())
            rels.addDocumentRelationship(pair.first, pair.second);

        QByteArray data;
        QBuffer buffer(&data);
        buffer.open(QIODevice::WriteOnly);
        rels.saveToXmlFile(&buffer);
        zipWriter.addFile(QStringLiteral("xl/drawings/_rels/drawing%1.xml.rels").arg(index), data);
        index += 1;
    }
}

void Package::writeImageFiles(ZipWriter &zipWriter)
{
    for (int i=0; i<m_workbook->images().size(); ++i) {
        QImage image = m_workbook->images()[i];

        QByteArray data;
        QBuffer buffer(&data);
        buffer.open(QIODevice::WriteOnly);
        image.save(&buffer, "png");
        zipWriter.addFile(QStringLiteral("xl/media/image%1.png").arg(i+1), data);
    }
}

} // namespace QXlsx
