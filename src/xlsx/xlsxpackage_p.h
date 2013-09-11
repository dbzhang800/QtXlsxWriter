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
#ifndef QXLSX_PACKAGE_H
#define QXLSX_PACKAGE_H

#include "xlsxglobal.h"
#include <QString>
class QIODevice;

namespace QXlsx {

class Workbook;
class ZipWriter;
class ZipReader;
class Document;
class Relationships;
class DocPropsCore;
class DocPropsApp;

class XLSX_AUTOTEST_EXPORT Package
{
public:
    Package(Document *document);

    bool parsePackage(QIODevice *packageDevice);
    bool createPackage(QIODevice *package);

private:

    void writeWorksheetFiles(ZipWriter &zipWriter);
//    void writeChartsheetFiles(ZipWriter &zipWriter);
    void writeWorkbookFile(ZipWriter &zipWriter);
//    void writeChartFiles(ZipWriter &zipWriter);
    void writeDrawingFiles(ZipWriter &zipWriter);
//    void writeVmlFiles(ZipWriter &zipWriter);
//    void writeCommentFiles(ZipWriter &zipWriter);
//    void writeTableFiles(ZipWriter &zipWriter);
    void writeSharedStringsFile(ZipWriter &zipWriter);
    void writeDocPropsAppFile(ZipWriter &zipWriter);
    void writeDocPropsCoreFile(ZipWriter &zipWriter);
    void writeContentTypesFile(ZipWriter &zipWriter);
    void writeStylesFiles(ZipWriter &zipWriter);
    void writeThemeFile(ZipWriter &zipWriter);
    void writeRootRelsFile(ZipWriter &zipWriter);
    void writeWorkbookRelsFile(ZipWriter &zipWriter);
    void writeWorksheetRelsFiles(ZipWriter &zipWriter);
//    void writeChartsheetRelsFile(ZipWriter &zipWriter);
    void writeDrawingRelsFiles(ZipWriter &zipWriter);
    void writeImageFiles(ZipWriter &zipWriter);
//    void writeVbaProjectFiles(ZipWriter &zipWriter);

    Document *m_document;
    Workbook *m_workbook;
    int m_worksheet_count;
    int m_chartsheet_count;
};

} // namespace QXlsx

#endif // QXLSX_PACKAGE_H
