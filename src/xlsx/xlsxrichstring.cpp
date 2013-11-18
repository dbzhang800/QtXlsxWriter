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
#include "xlsxrichstring_p.h"

QT_BEGIN_NAMESPACE_XLSX

RichString::RichString()
    :m_dirty(true)
{
}

RichString::RichString(const QString text)
    :m_dirty(true)
{
    addFragment(text, 0);
}

bool RichString::isRichString() const
{
    if (fragmentCount() > 1) //Is this enough??
        return true;
    return false;
}

bool RichString::isEmtpy() const
{
    return m_fragmentTexts.size() == 0;
}

QString RichString::toPlainString() const
{
    if (isEmtpy())
        return QString();
    if (m_fragmentTexts.size() == 1)
        return m_fragmentTexts[0];

    return m_fragmentTexts.join(QString());
}

int RichString::fragmentCount() const
{
    return m_fragmentTexts.size();
}

void RichString::addFragment(const QString &text, Format *format)
{
    m_fragmentTexts.append(text);
    m_fragmentFormats.append(format);
    m_dirty = true;
}

QString RichString::fragmentText(int index) const
{
    if (index < 0 || index >= fragmentCount())
        return QString();

    return m_fragmentTexts[index];
}

Format *RichString::fragmentFormat(int index) const
{
    if (index < 0 || index >= fragmentCount())
        return 0;

    return m_fragmentFormats[index];
}

/*!
 * \internal
 */
QByteArray RichString::idKey() const
{
    if (m_dirty) {
        RichString *rs = const_cast<RichString *>(this);
        QByteArray bytes;
        if (!isRichString()) {
            bytes = toPlainString().toUtf8();
        } else {
            //Generate a hash value base on QByteArray ?
            bytes.append("@@QtXlsxRichString=");
            for (int i=0; i<fragmentCount(); ++i) {
                bytes.append("@Text");
                bytes.append(m_fragmentTexts[i].toUtf8());
                bytes.append("@Format");
                if (m_fragmentFormats[i])
                    bytes.append(m_fragmentFormats[i]->fontKey());
            }
        }
        rs->m_idKey = bytes;
        rs->m_dirty = false;
    }

    return m_idKey;
}

bool operator==(const RichString &rs1, const RichString &rs2)
{
    if (rs1.fragmentCount() != rs2.fragmentCount())
        return false;

    return rs1.idKey() == rs2.idKey();
}

bool operator!=(const RichString &rs1, const RichString &rs2)
{
    if (rs1.fragmentCount() != rs2.fragmentCount())
        return true;

    return rs1.idKey() != rs2.idKey();
}

/*!
 * \internal
 */
bool operator<(const RichString &rs1, const RichString &rs2)
{
    return rs1.idKey() < rs2.idKey();
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
    return qHash(rs.idKey(), seed);
}

QT_END_NAMESPACE_XLSX
