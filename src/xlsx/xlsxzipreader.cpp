#include "xlsxzipreader_p.h"

#include <private/qzipreader_p.h>

namespace QXlsx {

ZipReader::ZipReader(const QString &filePath) :
    m_reader(new QZipReader(filePath))
{
    init();
}

ZipReader::ZipReader(QIODevice *device) :
    m_reader(new QZipReader(device))
{
    init();
}

ZipReader::~ZipReader()
{

}

void ZipReader::init()
{
    QList<QZipReader::FileInfo> allFiles = m_reader->fileInfoList();
    foreach (const QZipReader::FileInfo &fi, allFiles) {
        if (fi.isFile)
            m_filePaths.append(fi.filePath);
    }
}

bool ZipReader::exists() const
{
    return m_reader->exists();
}

QStringList ZipReader::filePaths() const
{
    return m_filePaths;
}

QByteArray ZipReader::fileData(const QString &fileName) const
{
    return m_reader->fileData(fileName);
}

} // namespace QXlsx
