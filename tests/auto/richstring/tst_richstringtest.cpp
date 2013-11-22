#include "xlsxrichstring.h"
#include <QtTest>
#include <QDebug>

class RichstringTest : public QObject
{
    Q_OBJECT

public:
    RichstringTest();

private Q_SLOTS:
    void testEqual();
};

RichstringTest::RichstringTest()
{
}

void RichstringTest::testEqual()
{
    QXlsx::RichString rs;
    rs.addFragment("Hello", QXlsx::Format());
    rs.addFragment(" RichText", QXlsx::Format());

    QXlsx::RichString rs2;
    rs2.addFragment("Hello", QXlsx::Format());
    rs2.addFragment(" Qt!", QXlsx::Format());

    QXlsx::RichString rs3;
    rs3.addFragment("Hello", QXlsx::Format());
    rs3.addFragment(" Qt!", QXlsx::Format());

    QVERIFY2(rs2 != rs, "Failure");
    QVERIFY2(rs2 == rs3, "Failure");
    QVERIFY2(rs2 != QStringLiteral("Hello Qt!"), "Failure");
}

QTEST_APPLESS_MAIN(RichstringTest)

#include "tst_richstringtest.moc"
