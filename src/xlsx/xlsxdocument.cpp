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

#include "xlsxdocument.h"
#include "xlsxdocument_p.h"
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"
#include "xlsxcontenttypes_p.h"
#include "xlsxrelationships_p.h"
#include "xlsxstyles_p.h"
#include "xlsxtheme_p.h"
#include "xlsxdocpropsapp_p.h"
#include "xlsxdocpropscore_p.h"
#include "xlsxsharedstrings_p.h"
#include "xlsxutility_p.h"
#include "xlsxworkbook_p.h"
#include "xlsxdrawing_p.h"
#include "xlsxmediafile_p.h"
#include "xlsxchart.h"
#include "xlsxzipreader_p.h"
#include "xlsxzipwriter_p.h"

#include <QFile>
#include <QPointF>
#include <QBuffer>
#include <QDir>

QT_BEGIN_NAMESPACE_XLSX

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

DocumentPrivate::DocumentPrivate(Document *p) :
    q_ptr(p), defaultPackageName(QStringLiteral("Book1.xlsx"))
{
    workbook = QSharedPointer<Workbook>(new Workbook);
}

void DocumentPrivate::init()
{
    if (workbook->worksheetCount() == 0)
        workbook->addWorksheet();
}

bool DocumentPrivate::loadPackage(QIODevice *device)
{
    Q_Q(Document);
    ZipReader zipReader(device);
    QStringList filePaths = zipReader.filePaths();

    //Load the Content_Types file
    if (!filePaths.contains(QLatin1String("[Content_Types].xml")))
        return false;
    contentTypes.loadFromXmlData(zipReader.fileData(QStringLiteral("[Content_Types].xml")));

    //Load root rels file
    if (!filePaths.contains(QLatin1String("_rels/.rels")))
        return false;
    Relationships rootRels;
    rootRels.loadFromXmlData(zipReader.fileData(QStringLiteral("_rels/.rels")));

    //load core property
    QList<XlsxRelationship> rels_core = rootRels.packageRelationships(QStringLiteral("/metadata/core-properties"));
    if (!rels_core.isEmpty()) {
        //Get the core property file name if it exists.
        //In normal case, this should be "docProps/core.xml"
        QString docPropsCore_Name = rels_core[0].target;

        DocPropsCore props;
        props.loadFromXmlData(zipReader.fileData(docPropsCore_Name));
        foreach (QString name, props.propertyNames())
            q->setDocumentProperty(name, props.property(name));
    }

    //load app property
    QList<XlsxRelationship> rels_app = rootRels.documentRelationships(QStringLiteral("/extended-properties"));
    if (!rels_app.isEmpty()) {
        //Get the app property file name if it exists.
        //In normal case, this should be "docProps/app.xml"
        QString docPropsApp_Name = rels_app[0].target;

        DocPropsApp props;
        props.loadFromXmlData(zipReader.fileData(docPropsApp_Name));
        foreach (QString name, props.propertyNames())
            q->setDocumentProperty(name, props.property(name));
    }

    //load workbook now, Get the workbook file path from the root rels file
    //In normal case, this should be "xl/workbook.xml"
    QList<XlsxRelationship> rels_xl = rootRels.documentRelationships(QStringLiteral("/officeDocument"));
    if (rels_xl.isEmpty())
        return false;
    QString xlworkbook_Path = rels_xl[0].target;
    QString xlworkbook_Dir = splitPath(xlworkbook_Path)[0];
    workbook->relationships().loadFromXmlData(zipReader.fileData(getRelFilePath(xlworkbook_Path)));
    workbook->setFilePath(xlworkbook_Path);
    workbook->loadFromXmlData(zipReader.fileData(xlworkbook_Path));

    //load styles
    QList<XlsxRelationship> rels_styles = workbook->relationships().documentRelationships(QStringLiteral("/styles"));
    if (!rels_styles.isEmpty()) {
        //In normal case this should be styles.xml which in xl
        QString name = rels_styles[0].target;
        QString path = xlworkbook_Dir + QLatin1String("/") + name;
        QSharedPointer<Styles> styles (new Styles(true));
        styles->loadFromXmlData(zipReader.fileData(path));
        workbook->d_func()->styles = styles;
    }

    //load sharedStrings
    QList<XlsxRelationship> rels_sharedStrings = workbook->relationships().documentRelationships(QStringLiteral("/sharedStrings"));
    if (!rels_sharedStrings.isEmpty()) {
        //In normal case this should be sharedStrings.xml which in xl
        QString name = rels_sharedStrings[0].target;
        QString path = xlworkbook_Dir + QLatin1String("/") + name;
        workbook->d_func()->sharedStrings->loadFromXmlData(zipReader.fileData(path));
    }

    //load theme
    QList<XlsxRelationship> rels_theme = workbook->relationships().documentRelationships(QStringLiteral("/theme"));
    if (!rels_theme.isEmpty()) {
        //In normal case this should be theme/theme1.xml which in xl
        QString name = rels_theme[0].target;
        QString path = xlworkbook_Dir + QLatin1String("/") + name;
        workbook->theme()->loadFromXmlData(zipReader.fileData(path));
    }

    //load worksheets
    for (int i=0; i<workbook->worksheetCount(); ++i) {
        Worksheet *sheet = workbook->worksheet(i);
        QString rel_path = getRelFilePath(sheet->filePath());
        //If the .rel file exists, load it.
        if (zipReader.filePaths().contains(rel_path))
            sheet->relationships().loadFromXmlData(zipReader.fileData(rel_path));
        sheet->loadFromXmlData(zipReader.fileData(sheet->filePath()));
    }

    //load drawings
    for (int i=0; i<workbook->drawings().size(); ++i) {
        Drawing *drawing = workbook->drawings()[i];
        QString rel_path = getRelFilePath(drawing->filePath());
        if (zipReader.filePaths().contains(rel_path))
            drawing->relationships.loadFromXmlData(zipReader.fileData(rel_path));
        drawing->loadFromXmlData(zipReader.fileData(drawing->filePath()));
    }

    //load charts
    QList<QSharedPointer<Chart> > chartFileToLoad = workbook->chartFiles();
    for (int i=0; i<chartFileToLoad.size(); ++i) {
        QSharedPointer<Chart> cf = chartFileToLoad[i];
        cf->loadFromXmlData(zipReader.fileData(cf->filePath()));
    }

    //load media files
    QList<QSharedPointer<MediaFile> > mediaFileToLoad = workbook->mediaFiles();
    for (int i=0; i<mediaFileToLoad.size(); ++i) {
        QSharedPointer<MediaFile> mf = mediaFileToLoad[i];
        const QString path = mf->fileName();
        const QString suffix = path.mid(path.lastIndexOf(QLatin1Char('.'))+1);
        mf->set(zipReader.fileData(path), suffix);
    }

    return true;
}

bool DocumentPrivate::savePackage(QIODevice *device) const
{
    Q_Q(const Document);
    ZipWriter zipWriter(device);
    if (zipWriter.error())
        return false;

    contentTypes.clearOverrides();

    DocPropsApp docPropsApp;
    DocPropsCore docPropsCore;

    // save worksheet xml files
    for (int i=0; i<workbook->worksheetCount(); ++i) {
        Worksheet *sheet = workbook->worksheet(i);
        contentTypes.addWorksheetName(QStringLiteral("sheet%1").arg(i+1));
        docPropsApp.addPartTitle(sheet->sheetName());

        zipWriter.addFile(QStringLiteral("xl/worksheets/sheet%1.xml").arg(i+1), sheet->saveToXmlData());
        Relationships &rel = sheet->relationships();
        if (!rel.isEmpty())
            zipWriter.addFile(QStringLiteral("xl/worksheets/_rels/sheet%1.xml.rels").arg(i+1), rel.saveToXmlData());
    }

    // save workbook xml file
    contentTypes.addWorkbook();
    zipWriter.addFile(QStringLiteral("xl/workbook.xml"), workbook->saveToXmlData());
    zipWriter.addFile(QStringLiteral("xl/_rels/workbook.xml.rels"), workbook->relationships().saveToXmlData());

    // save drawing xml files
    for (int i=0; i<workbook->drawings().size(); ++i) {
        contentTypes.addDrawingName(QStringLiteral("drawing%1").arg(i+1));

        Drawing *drawing = workbook->drawings()[i];
        zipWriter.addFile(QStringLiteral("xl/drawings/drawing%1.xml").arg(i+1), drawing->saveToXmlData());
        if (!drawing->relationships.isEmpty())
            zipWriter.addFile(QStringLiteral("xl/drawings/_rels/drawing%1.xml.rels").arg(i+1), drawing->relationships.saveToXmlData());
    }

    // save docProps app/core xml file
    foreach (QString name, q->documentPropertyNames()) {
        docPropsApp.setProperty(name, q->documentProperty(name));
        docPropsCore.setProperty(name, q->documentProperty(name));
    }
    if (workbook->worksheetCount())
        docPropsApp.addHeadingPair(QStringLiteral("Worksheets"), workbook->worksheetCount());
    contentTypes.addDocPropApp();
    contentTypes.addDocPropCore();
    zipWriter.addFile(QStringLiteral("docProps/app.xml"), docPropsApp.saveToXmlData());
    zipWriter.addFile(QStringLiteral("docProps/core.xml"), docPropsCore.saveToXmlData());

    // save sharedStrings xml file
    if (!workbook->sharedStrings()->isEmpty()) {
        contentTypes.addSharedString();
        zipWriter.addFile(QStringLiteral("xl/sharedStrings.xml"), workbook->sharedStrings()->saveToXmlData());
    }

    // save styles xml file
    contentTypes.addStyles();
    zipWriter.addFile(QStringLiteral("xl/styles.xml"), workbook->styles()->saveToXmlData());

    // save theme xml file
    contentTypes.addTheme();
    zipWriter.addFile(QStringLiteral("xl/theme/theme1.xml"), workbook->theme()->saveToXmlData());

    // save chart xml files
    for (int i=0; i<workbook->chartFiles().size(); ++i) {
        contentTypes.addChartName(QStringLiteral("chart%1").arg(i+1));
        QSharedPointer<Chart> cf = workbook->chartFiles()[i];
        zipWriter.addFile(QStringLiteral("xl/charts/chart%1.xml").arg(i+1), cf->saveToXmlData());
    }

    // save image files
    for (int i=0; i<workbook->mediaFiles().size(); ++i) {
        QSharedPointer<MediaFile> mf = workbook->mediaFiles()[i];
        if (!mf->mimeType().isEmpty())
            contentTypes.addDefault(mf->suffix(), mf->mimeType());

        zipWriter.addFile(QStringLiteral("xl/media/image%1.%2").arg(i+1).arg(mf->suffix()), mf->contents());
    }

    // save root .rels xml file
    Relationships rootrels;
    rootrels.addDocumentRelationship(QStringLiteral("/officeDocument"), QStringLiteral("xl/workbook.xml"));
    rootrels.addPackageRelationship(QStringLiteral("/metadata/core-properties"), QStringLiteral("docProps/core.xml"));
    rootrels.addDocumentRelationship(QStringLiteral("/extended-properties"), QStringLiteral("docProps/app.xml"));
    zipWriter.addFile(QStringLiteral("_rels/.rels"), rootrels.saveToXmlData());

    // save content types xml file
    zipWriter.addFile(QStringLiteral("[Content_Types].xml"), contentTypes.saveToXmlData());

    zipWriter.close();
    return true;
}


/*!
  \class Document
  \inmodule QtXlsx
  \brief The Document class provides a API that is used to handle the contents of .xlsx files.

*/

/*!
 * Creates a new empty xlsx document.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    d_ptr->init();
}

/*!
 * \overload
 * Try to open an existing xlsx document named \a name.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(const QString &name, QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    d_ptr->packageName = name;
    if (QFile::exists(name)) {
        QFile xlsx(name);
        if (xlsx.open(QFile::ReadOnly))
            d_ptr->loadPackage(&xlsx);
    }
    d_ptr->init();
}

/*!
 * \overload
 * Try to open an existing xlsx document from \a device.
 * The \a parent argument is passed to QObject's constructor.
 */
Document::Document(QIODevice *device, QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    if (device && device->isReadable())
        d_ptr->loadPackage(device);
    d_ptr->init();
}

/*!
    \overload

    Write \a value to cell \a row_column with the \a format.
 */
int Document::write(const QString &row_column, const QVariant &value, const Format &format)
{
    return currentWorksheet()->write(row_column, value, format);
}

/*!
 * Write \a value to cell (\a row, \a col) with the \a format.
 */
int Document::write(int row, int col, const QVariant &value, const Format &format)
{
    return currentWorksheet()->write(row, col, value, format);
}

/*!
    \overload
    Returns the contents of the cell \a cell.
*/
QVariant Document::read(const QString &cell) const
{
    return currentWorksheet()->read(cell);
}

/*!
    Returns the contents of the cell (\a row, \a col).
 */
QVariant Document::read(int row, int col) const
{
    return currentWorksheet()->read(row, col);
}

/*!
 * Insert an \a image to current active worksheet at the position \a row, \a column
 * Returns ture if success.
 */
bool Document::insertImage(int row, int column, const QImage &image)
{
    return currentWorksheet()->insertImage(row, column, image);
}

/*!
 * Creates an chart with the given \a size and insert it to the current
 * active worksheet at the position \a row, \a col.
 * The chart will be returned.
 */
Chart *Document::insertChart(int row, int col, const QSize &size)
{
    return currentWorksheet()->insertChart(row, col, size);
}

/*!
 * \overload
 * \deprecated
 * Insert an \a image to current active worksheet to the position \a row, \a column with the given
 * \a xOffset, \a yOffset, \a xScale and \a yScale.
 */
int Document::insertImage(int row, int column, const QImage &image, double /*xOffset*/, double /*yOffset*/, double /*xScale*/, double /*yScale*/)
{
    return currentWorksheet()->insertImage(row, column, image);
}

/*!
    Merge a \a range of cells. The first cell should contain the data and the others should
    be blank. All cells will be applied the same style if a valid \a format is given.

    \note All cells except the top-left one will be cleared.
 */
int Document::mergeCells(const CellRange &range, const Format &format)
{
    return currentWorksheet()->mergeCells(range, format);
}

/*!
    \overload
    Merge a \a range of cells. The first cell should contain the data and the others should
    be blank. All cells will be applied the same style if a valid \a format is given.

    \note All cells except the top-left one will be cleared.
 */
int Document::mergeCells(const QString &range, const Format &format)
{
    return currentWorksheet()->mergeCells(range, format);
}

/*!
    Unmerge the cells in the \a range.
*/
int Document::unmergeCells(const QString &range)
{
    return currentWorksheet()->unmergeCells(range);
}

/*!
    Unmerge the cells in the \a range.
*/
int Document::unmergeCells(const CellRange &range)
{
    return currentWorksheet()->unmergeCells(range);
}

/*!
  Sets the properties of \a row with the given \a height, \a format and \a hidden.
  \a row is 1-indexed.

  Returns false if failed.
 */
bool Document::setRow(int row, double height, const Format &format, bool hidden)
{
    return currentWorksheet()->setRow(row, height, format, hidden);
}

/*!
  Sets the column properties for all columns from \a colFirst to \a colLast with
  the given \a width, \a format and \a hidden. Column
  width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font.
  \a colFirst and \a colLast are all 1-indexed.

  Return false if failed.
 */
bool Document::setColumn(int colFirst, int colLast, double width, const Format &format, bool hidden)
{
    return currentWorksheet()->setColumn(colFirst, colLast, width, format, hidden);
}

/*!
  \overload

  Sets column width and format for all columns from \a colFirst to \a colLast with
  the given \a width and \a format. Column
  \a width measured as the number of characters of the maximum digit width of the
  numbers 0, 1, 2, ..., 9 as rendered in the normal style's font.
  \a colFirst and \a colLast should be "A", "B", "C", ...
 */
bool Document::setColumn(const QString &colFirst, const QString &colLast, double width, const Format &format, bool hidden)
{
    return currentWorksheet()->setColumn(colFirst, colLast, width, format, hidden);
}

/*!
   Groups rows from \a rowFirst to \a rowLast with the given \a collapsed.
   Returns false if error occurs.
 */
bool Document::groupRows(int rowFirst, int rowLast, bool collapsed)
{
    return currentWorksheet()->groupRows(rowFirst, rowLast, collapsed);
}

/*!
   Groups columns from \a colFirst to \a colLast with the given \a collapsed.
   Returns false if error occurs.
 */
bool Document::groupColumns(int colFirst, int colLast, bool collapsed)
{
    return currentWorksheet()->groupColumns(colFirst, colLast, collapsed);
}

/*!
 *  Add a data \a validation rule for current worksheet. Returns true if successful.
 */
bool Document::addDataValidation(const DataValidation &validation)
{
    return currentWorksheet()->addDataValidation(validation);
}

/*!
 *  Add a  conditional formatting \a cf for current worksheet. Returns true if successful.
 */
bool Document::addConditionalFormatting(const ConditionalFormatting &cf)
{
    return currentWorksheet()->addConditionalFormatting(cf);
}

/*!
 * Returns a Cell object based on the given \a pos. 0 will be returned if the cell doesn't exist.
 */
Cell *Document::cellAt(const QString &pos) const
{
    return currentWorksheet()->cellAt(pos);
}

/*!
 * Returns a Cell object based on the given \a row and \a col.
 */
Cell *Document::cellAt(int row, int col) const
{
    return currentWorksheet()->cellAt(row, col);
}

/*!
 * \brief Create a defined name in the workbook with the given \a name, \a formula, \a comment
 *  and \a scope.
 *
 * \param name The defined name.
 * \param formula The cell or range that the defined name refers to.
 * \param scope The name of one worksheet, or empty which means golbal scope.
 * \return Return false if the name invalid.
 */
bool Document::defineName(const QString &name, const QString &formula, const QString &comment, const QString &scope)
{
    Q_D(Document);

    return d->workbook->defineName(name, formula, comment, scope);
}

/*!
    Return the range that contains cell data.
 */
CellRange Document::dimension() const
{
    return currentWorksheet()->dimension();
}

/*!
 * Returns the value of the document's \a key property.
 */
QString Document::documentProperty(const QString &key) const
{
    Q_D(const Document);
    if (d->documentProperties.contains(key))
        return d->documentProperties[key];
    else
        return QString();
}

/*!
    Set the document properties such as Title, Author etc.

    The method can be used to set the document properties of the Excel
    file created by Qt Xlsx. These properties are visible when you use the
    Office Button -> Prepare -> Properties option in Excel and are also
    available to external applications that read or index windows files.

    The \a property \a key that can be set are:

    \list
    \li title
    \li subject
    \li creator
    \li manager
    \li company
    \li category
    \li keywords
    \li description
    \li status
    \endlist
*/
void Document::setDocumentProperty(const QString &key, const QString &property)
{
    Q_D(Document);
    d->documentProperties[key] = property;
}

/*!
 * Returns the names of all properties that were addedusing setDocumentProperty().
 */
QStringList Document::documentPropertyNames() const
{
    Q_D(const Document);
    return d->documentProperties.keys();
}

/*!
 * Return the internal Workbook object.
 */
Workbook *Document::workbook() const
{
    Q_D(const Document);
    return d->workbook.data();
}

/*!
 * Returns the worksheet object named \a sheetName.
 */
Worksheet *Document::worksheet(const QString &sheetName) const
{
    Q_D(const Document);
    return d->workbook->worksheet(worksheetNames().indexOf(sheetName));
}

/*!
 * Creates and append an document with name \a name.
 * Return true if success.
 */
bool Document::addWorksheet(const QString &name)
{
    Q_D(Document);
    return d->workbook->addWorksheet(name);
}

/*!
 * Creates and inserts an document with name \a name at the \a index.
 * Returns false if the \a name already used.
 */
bool Document::insertWorkSheet(int index, const QString &name)
{
    Q_D(Document);
    return d->workbook->insertWorkSheet(index, name);
}

/*!
   Rename the worksheet from \a oldName to \a newName.
   Returns true if the success.
 */
bool Document::renameWorksheet(const QString &oldName, const QString &newName)
{
    Q_D(Document);
    if (oldName == newName)
        return false;
    return d->workbook->renameWorksheet(worksheetNames().indexOf(oldName), newName);
}

/*!
   Make a copy of the worksheet \a srcName with the new name \a distName.
   Returns true if the success.
 */
bool Document::copyWorksheet(const QString &srcName, const QString &distName)
{
    Q_D(Document);
    if (srcName == distName)
        return false;
    return d->workbook->copyWorksheet(worksheetNames().indexOf(srcName), distName);
}

/*!
   Move the worksheet \a srcName to the new pos \a distIndex.
   Returns true if the success.
 */
bool Document::moveWorksheet(const QString &srcName, int distIndex)
{
    Q_D(Document);
    return d->workbook->moveWorksheet(worksheetNames().indexOf(srcName), distIndex);
}

/*!
   Delete the worksheet \a name.
   Returns true if current sheet was deleted successfully.
 */
bool Document::deleteWorksheet(const QString &name)
{
    Q_D(Document);
    return d->workbook->deleteWorksheet(worksheetNames().indexOf(name));
}

/*!
   \deprecated
   Rename current worksheet to new \a name.
   Returns true if the name defined successful.

   \sa renameWorksheet()
 */
bool Document::setSheetName(const QString &name)
{
    return renameWorksheet(currentWorksheet()->sheetName(), name);
}

/*!
 * \brief Return pointer of current worksheet.
 */
Worksheet *Document::currentWorksheet() const
{
    Q_D(const Document);
    if (d->workbook->worksheetCount() == 0)
        return 0;

    return d->workbook->activeWorksheet();
}

/*!
 *  \deprecated
 *  Set current worksheet to be the sheet at \a index.
 *  \sa selectWorksheet()
 */
void Document::setCurrentWorksheet(int index)
{
    Q_D(Document);
    d->workbook->setActiveWorksheet(index);
}

/*!
 *  \deprecated
 *  Set current selected worksheet to be the sheet named \a name.
 *  \sa selectWorksheet()
 */
void Document::setCurrentWorksheet(const QString &name)
{
    selectWorksheet(name);
}

/*!
 * \brief Set worksheet named \a name to be active sheet.
 * Returns true if success.
 */
bool Document::selectWorksheet(const QString &name)
{
    Q_D(Document);
    return d->workbook->setActiveWorksheet(worksheetNames().indexOf(name));
}

/*!
 * Returns the names of worksheets contained in current document.
 */
QStringList Document::worksheetNames() const
{
    Q_D(const Document);
    return d->workbook->worksheetNames();
}

/*!
 * Save current document to the filesystem. If no name specified when
 * the document constructed, a default name "book1.xlsx" will be used.
 * Returns true if saved successfully.
 */
bool Document::save() const
{
    Q_D(const Document);
    QString name = d->packageName.isEmpty() ? d->defaultPackageName : d->packageName;

    return saveAs(name);
}

/*!
 * Saves the document to the file with the given \a name.
 * Returns true if saved successfully.
 */
bool Document::saveAs(const QString &name) const
{
    QFile file(name);
    if (file.open(QIODevice::WriteOnly))
        return saveAs(&file);
    return false;
}

/*!
 * \overload
 * This function writes a document to the given \a device.
 */
bool Document::saveAs(QIODevice *device) const
{
    Q_D(const Document);
    return d->savePackage(device);
}

/*!
 * Destroys the document and cleans up.
 */
Document::~Document()
{
    delete d_ptr;
}

QT_END_NAMESPACE_XLSX
