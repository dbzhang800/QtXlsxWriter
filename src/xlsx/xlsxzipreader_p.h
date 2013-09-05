#ifndef QXLSX_XLSXZIPREADER_P_H
#define QXLSX_XLSXZIPREADER_P_H

#include "xlsxglobal.h"
#include <QScopedPointer>
#include <QStringList>
class QZipReader;
class QIODevice;

namespace QXlsx {

class XLSX_AUTOTEST_EXPORT ZipReader
{
public:
    explicit ZipReader(const QString &fileName);
    explicit ZipReader(QIODevice *device);
    ~ZipReader();
    bool exists() const;
    QStringList filePaths() const;
    QByteArray fileData(const QString &fileName) const;

private:
    Q_DISABLE_COPY(ZipReader)
    void init();
    QScopedPointer<QZipReader> m_reader;
    QStringList m_filePaths;
};

} // namespace QXlsx

#endif // QXLSX_XLSXZIPREADER_P_H
