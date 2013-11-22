/****************************************************************************
** Copyright (c) 2013 Debao Zhang <hello@debao.me>
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
#include "xlsxrichstring.h"
#include "xlsxrichstring_p.h"
#include "xlsxformat_p.h"

QT_BEGIN_NAMESPACE_XLSX

RichStringPrivate::RichStringPrivate()
    :_dirty(true)
{

}

RichStringPrivate::RichStringPrivate(const RichStringPrivate &other)
    :QSharedData(other), fragmentTexts(other.fragmentTexts)
    ,fragmentFormats(other.fragmentFormats)
    , _idKey(other.idKey()), _dirty(other._dirty)
{

}

RichStringPrivate::~RichStringPrivate()
{

}

/*!
    \class RichString

*/

RichString::RichString()
    :d(new RichStringPrivate)
{
}

RichString::RichString(const QString text)
    :d(new RichStringPrivate)
{
    addFragment(text, Format());
}

RichString::RichString(const RichString &other)
    :d(other.d)
{

}

RichString::~RichString()
{

}

RichString &RichString::operator =(const RichString &other)
{
    this->d = other.d;
    return *this;
}

/*!
    Returns the rich string as a QVariant
*/
RichString::operator QVariant() const
{
    return QVariant(qMetaTypeId<RichString>(), this);
}

bool RichString::isRichString() const
{
    if (fragmentCount() > 1) //Is this enough??
        return true;
    return false;
}

bool RichString::isEmtpy() const
{
    return d->fragmentTexts.size() == 0;
}

QString RichString::toPlainString() const
{
    if (isEmtpy())
        return QString();
    if (d->fragmentTexts.size() == 1)
        return d->fragmentTexts[0];

    return d->fragmentTexts.join(QString());
}

int RichString::fragmentCount() const
{
    return d->fragmentTexts.size();
}

void RichString::addFragment(const QString &text, const Format &format)
{
    d->fragmentTexts.append(text);
    d->fragmentFormats.append(format);
    d->_dirty = true;
}

QString RichString::fragmentText(int index) const
{
    if (index < 0 || index >= fragmentCount())
        return QString();

    return d->fragmentTexts[index];
}

Format RichString::fragmentFormat(int index) const
{
    if (index < 0 || index >= fragmentCount())
        return Format();

    return d->fragmentFormats[index];
}

/*!
 * \internal
 */
QByteArray RichStringPrivate::idKey() const
{
    if (_dirty) {
        RichStringPrivate *rs = const_cast<RichStringPrivate *>(this);
        QByteArray bytes;
        if (fragmentTexts.size() == 1) {
            bytes = fragmentTexts[0].toUtf8();
        } else {
            //Generate a hash value base on QByteArray ?
            bytes.append("@@QtXlsxRichString=");
            for (int i=0; i<fragmentTexts.size(); ++i) {
                bytes.append("@Text");
                bytes.append(fragmentTexts[i].toUtf8());
                bytes.append("@Format");
                if (fragmentFormats[i].hasFontData())
                    bytes.append(fragmentFormats[i].fontKey());
            }
        }
        rs->_idKey = bytes;
        rs->_dirty = false;
    }

    return _idKey;
}

bool operator==(const RichString &rs1, const RichString &rs2)
{
    if (rs1.fragmentCount() != rs2.fragmentCount())
        return false;

    return rs1.d->idKey() == rs2.d->idKey();
}

bool operator!=(const RichString &rs1, const RichString &rs2)
{
    if (rs1.fragmentCount() != rs2.fragmentCount())
        return true;

    return rs1.d->idKey() != rs2.d->idKey();
}

/*!
 * \internal
 */
bool operator<(const RichString &rs1, const RichString &rs2)
{
    return rs1.d->idKey() < rs2.d->idKey();
}

bool operator ==(const RichString &rs1, const QString &rs2)
{
    if (rs1.fragmentCount() == 1 && rs1.fragmentText(0) == rs2) //format == 0
        return true;

    return false;
}

bool operator !=(const RichString &rs1, const QString &rs2)
{
    if (rs1.fragmentCount() == 1 && rs1.fragmentText(0) == rs2) //format == 0
        return false;

    return true;
}

bool operator ==(const QString &rs1, const RichString &rs2)
{
    return rs2 == rs1;
}

bool operator !=(const QString &rs1, const RichString &rs2)
{
    return rs2 != rs1;
}

uint qHash(const RichString &rs, uint seed) Q_DECL_NOTHROW
{
    return qHash(rs.d->idKey(), seed);
}

QT_END_NAMESPACE_XLSX
