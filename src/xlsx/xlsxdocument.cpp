#include "xlsxdocument.h"
#include "xlsxdocument_p.h"
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"
#include "xlsxpackage_p.h"

#include <QFile>
#include <QPointF>

namespace QXlsx {

DocumentPrivate::DocumentPrivate(Document *p) :
    q_ptr(p), defaultPackageName(QStringLiteral("Book1.xlsx"))
{
    workbook = new Workbook(p);
}

void DocumentPrivate::init()
{
    if (workbook->worksheets().size() == 0)
        workbook->addWorksheet();
}

bool DocumentPrivate::loadPackage(QIODevice *device)
{
    Q_Q(Document);
    Package package(q);
    return package.parsePackage(device);
}


/*!
  \class Document

*/

Document::Document(QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    d_ptr->init();
}

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

Document::Document(QIODevice *device, QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    if (device && device->isReadable())
        d_ptr->loadPackage(device);
    d_ptr->init();
}

Format *Document::createFormat()
{
    Q_D(Document);
    return d->workbook->createFormat();
}

int Document::write(const QString row_column, const QVariant &value, Format *format)
{
    return activedWorksheet()->write(row_column, value, format);
}

int Document::write(int row, int col, const QVariant &value, Format *format)
{
    return activedWorksheet()->write(row, col, value, format);
}

int Document::insertImage(int row, int column, const QImage &image, double xOffset, double yOffset, double xScale, double yScale)
{
    return activedWorksheet()->insertImage(row, column, image, QPointF(xOffset, yOffset), xScale, yScale);
}

int Document::mergeCells(const QString &range)
{
    return activedWorksheet()->mergeCells(range);
}

int Document::unmergeCells(const QString &range)
{
    return activedWorksheet()->unmergeCells(range);
}

bool Document::setRow(int row, double height, Format *format, bool hidden)
{
    return activedWorksheet()->setRow(row, height, format, hidden);
}

bool Document::setColumn(int colFirst, int colLast, double width, Format *format, bool hidden)
{
    return activedWorksheet()->setColumn(colFirst, colLast, width, format, hidden);
}

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

    The properties \a key that can be set are:

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

QStringList Document::documentPropertyNames() const
{
    Q_D(const Document);
    return d->documentProperties.keys();
}

Workbook *Document::workbook() const
{
    Q_D(const Document);
    return d->workbook;
}

bool Document::addWorksheet(const QString &name)
{
    Q_D(Document);
    return d->workbook->addWorksheet(name);
}

bool Document::insertWorkSheet(int index, const QString &name)
{
    Q_D(Document);
    return d->workbook->insertWorkSheet(index, name);
}

Worksheet *Document::activedWorksheet() const
{
    Q_D(const Document);
    if (d->workbook->worksheets().size() == 0)
        return 0;

    return d->workbook->worksheets().at(d->workbook->activedWorksheet());
}

int Document::activedWorksheetIndex() const
{
    Q_D(const Document);
    return d->workbook->activedWorksheet();
}

void Document::setActivedWorksheetIndex(int index)
{
    Q_D(Document);
    d->workbook->setActivedWorksheet(index);
}

bool Document::save()
{
    Q_D(Document);
    QString name = d->packageName.isEmpty() ? d->defaultPackageName : d->packageName;

    return saveAs(name);
}

bool Document::saveAs(const QString &name)
{
    QFile file(name);
    if (file.open(QIODevice::WriteOnly))
        return saveAs(&file);
    return false;
}

bool Document::saveAs(QIODevice *device)
{
    Q_D(Document);

//    activedWorksheet()->setHidden(false);
//    activedWorksheet()->setSelected(true);

    //Create the package based on current workbook
    Package package(this);
    return package.createPackage(device);
}

Document::~Document()
{
    delete d_ptr;
}

} // namespace QXlsx
