#include "xlsxdocument.h"
#include "xlsxdocument_p.h"
#include "xlsxworkbook.h"
#include "xlsxworksheet.h"

#include <QFile>

namespace QXlsx {

DocumentPrivate::DocumentPrivate(Document *p) :
    q_ptr(p), defaultPackageName(QStringLiteral("Book1"))
{

}

bool DocumentPrivate::loadPackage(QIODevice *device)
{

    return false;
}


/*!
  \class Document

*/

Document::Document(QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
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
}

Document::Document(QIODevice *device, QObject *parent) :
    QObject(parent), d_ptr(new DocumentPrivate(this))
{
    if (device && device->isReadable())
        d_ptr->loadPackage(device);
}

bool Document::save()
{
    return false;
}

bool Document::saveAs(const QString &name)
{
    return false;
}

bool Document::saveAs(QIODevice *device)
{
    return false;
}

Document::~Document()
{
    delete d_ptr;
}

} // namespace QXlsx
