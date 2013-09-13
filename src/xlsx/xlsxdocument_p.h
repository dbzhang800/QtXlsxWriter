#ifndef XLSXDOCUMENT_P_H
#define XLSXDOCUMENT_P_H

#include "xlsxdocument.h"
#include "xlsxworkbook.h"

#include <QMap>

namespace QXlsx {

class DocumentPrivate
{
    Q_DECLARE_PUBLIC(Document)
public:
    DocumentPrivate(Document *p);
    void init();

    bool loadPackage(QIODevice *device);

    Document *q_ptr;
    const QString defaultPackageName; //default name when package name not specified
    QString packageName; //name of the .xlsx file

    QMap<QString, QString> documentProperties; //core, app and custom properties
    QSharedPointer<Workbook> workbook;
};

}

#endif // XLSXDOCUMENT_P_H
