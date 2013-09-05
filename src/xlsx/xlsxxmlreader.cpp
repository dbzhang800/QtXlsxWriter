#include "xlsxxmlreader_p.h"

namespace QXlsx {

XmlStreamReader::XmlStreamReader(QIODevice *device) :
    QXmlStreamReader(device)
{
}

} // namespace QXlsx
