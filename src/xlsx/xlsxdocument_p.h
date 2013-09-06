#ifndef XLSXDOCUMENT_P_H
#define XLSXDOCUMENT_P_H

#include "xlsxdocument.h"

namespace QXlsx {

class DocumentPrivate
{
    Q_DECLARE_PUBLIC(Document)
public:
    DocumentPrivate(Document *p);

    bool loadPackage(QIODevice *device);

    Document *q_ptr;
    const QString defaultPackageName; //default name when package name not specified
    QString packageName; //name of the .xlsx file

};

}

#endif // XLSXDOCUMENT_P_H
