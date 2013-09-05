#ifndef QXLSX_XLSXXMLREADER_H
#define QXLSX_XLSXXMLREADER_H

#include <QXmlStreamReader>

namespace QXlsx {

class XmlStreamReader : public QXmlStreamReader
{
public:
    XmlStreamReader(QIODevice *device);
};

} // namespace QXlsx

#endif // QXLSX_XLSXXMLREADER_H
