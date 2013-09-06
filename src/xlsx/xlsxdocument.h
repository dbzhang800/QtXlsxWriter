#ifndef QXLSX_XLSXDOCUMENT_H
#define QXLSX_XLSXDOCUMENT_H

#include "xlsxglobal.h"
#include <QObject>
class QIODevice;

namespace QXlsx {

class Workbook;
class Worksheet;
class DocumentPrivate;
class Q_XLSX_EXPORT Document : public QObject
{
    Q_OBJECT
    Q_DECLARE_PRIVATE(Document)

public:
    explicit Document(QObject *parent = 0);
    Document(const QString &name, QObject *parent=0);
    Document(QIODevice *device, QObject *parent=0);
    ~Document();

    bool save();
    bool saveAs(const QString &name);
    bool saveAs(QIODevice *device);

private:
    Q_DISABLE_COPY(Document)
    DocumentPrivate * const d_ptr;
};

} // namespace QXlsx

#endif // QXLSX_XLSXDOCUMENT_H
