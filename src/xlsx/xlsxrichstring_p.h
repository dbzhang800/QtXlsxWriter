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
#ifndef XLSXRICHSTRING_P_H
#define XLSXRICHSTRING_P_H

#include "xlsxglobal.h"
#include "xlsxformat.h"
#include <QStringList>

QT_BEGIN_NAMESPACE_XLSX

class RichString;
// qHash is a friend, but we can't use default arguments for friends (§8.3.6.4)
XLSX_AUTOTEST_EXPORT uint qHash(const RichString &rs, uint seed = 0) Q_DECL_NOTHROW;

class XLSX_AUTOTEST_EXPORT RichString
{
public:
    RichString();
    explicit RichString(const QString text);
    bool isRichString() const;
    bool isEmtpy() const;
    QString toPlainString() const;

    int fragmentCount() const;
    void addFragment(const QString &text, Format *format);
    QString fragmentText(int index) const;
    Format *fragmentFormat(int index) const;

private:
    friend XLSX_AUTOTEST_EXPORT uint qHash(const RichString &rs, uint seed) Q_DECL_NOTHROW;
    friend XLSX_AUTOTEST_EXPORT bool operator==(const RichString &rs1, const RichString &rs2);
    friend XLSX_AUTOTEST_EXPORT bool operator!=(const RichString &rs1, const RichString &rs2);
    friend XLSX_AUTOTEST_EXPORT bool operator<(const RichString &rs1, const RichString &rs2);

    QByteArray idKey() const;

    QStringList m_fragmentTexts;
    QList<Format *> m_fragmentFormats;
    QByteArray m_idKey;
    bool m_dirty;
};


XLSX_AUTOTEST_EXPORT bool operator==(const RichString &rs1, const RichString &rs2);
XLSX_AUTOTEST_EXPORT bool operator!=(const RichString &rs1, const RichString &rs2);
XLSX_AUTOTEST_EXPORT bool operator<(const RichString &rs1, const RichString &rs2);
XLSX_AUTOTEST_EXPORT bool operator==(const RichString &rs1, const QString &rs2);
XLSX_AUTOTEST_EXPORT bool operator==(const QString &rs1, const RichString &rs2);
XLSX_AUTOTEST_EXPORT bool operator!=(const RichString &rs1, const QString &rs2);
XLSX_AUTOTEST_EXPORT bool operator!=(const QString &rs1, const RichString &rs2);

QT_END_NAMESPACE_XLSX

Q_DECLARE_METATYPE(QXlsx::RichString);

#endif // XLSXRICHSTRING_P_H
