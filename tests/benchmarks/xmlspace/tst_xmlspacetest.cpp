#include <QString>
#include <QRegularExpression>
#include <QtTest>

bool startsWithOrEndsWithSpace(const QString &s, int flag)
{
    if (flag == 0) {
        return (s.contains(QRegularExpression("^\\s")) || s.contains(QRegularExpression("\\s$")));
    } else if (flag == 1) {
        return (s.contains(QRegularExpression("^\\s|\\s$")));
    } else if (flag == 2) {
        static QRegularExpression re("^\\s|\\s$");
        return s.contains(re);
    } else if (flag == 3) {
        return s.startsWith(QLatin1Char(' ')) || s.endsWith(QLatin1Char(' '))
                || s.startsWith(QLatin1Char('\t')) || s.endsWith(QLatin1Char('\t'))
                || s.startsWith(QLatin1Char('\r')) || s.endsWith(QLatin1Char('\r'))
                || s.startsWith(QLatin1Char('\n')) || s.endsWith(QLatin1Char('\n'));
    } else if (flag == 4) {
        //static QString spaces(" \t\n\r");
        QString spaces(QStringLiteral(" \t\n\r"));
        return !s.isEmpty() && (spaces.contains(s.at(0))||spaces.contains(s.at(s.length()-1)));
    } else {
        return false;
    }
}

class XmlspaceTest : public QObject
{
    Q_OBJECT

public:
    XmlspaceTest();

private Q_SLOTS:
    void teststartsWithOrEndsWithSpace();
    void teststartsWithOrEndsWithSpace_data();

    void testCase1();
    void testCase1_data();
};

XmlspaceTest::XmlspaceTest()
{
}

void XmlspaceTest::teststartsWithOrEndsWithSpace()
{
    QFETCH(QString, data);
    QFETCH(bool, res);

    for (int f=0; f<5; ++f) {
        QCOMPARE(startsWithOrEndsWithSpace(data, f), res);
    }
}

void XmlspaceTest::teststartsWithOrEndsWithSpace_data()
{
    //QTest::addColumn<int>("flag");
    QTest::addColumn<QString>("data");
    QTest::addColumn<bool>("res");

    QTest::newRow("")<<QString()<<false;
    QTest::newRow("")<<""<<false;
    QTest::newRow("")<<"  "<<true;
    QTest::newRow("")<<"A  B"<<false;
    QTest::newRow("")<<" A  B"<<true;
    QTest::newRow("")<<"A  B\t"<<true;
    QTest::newRow("")<<" \tA  B\t"<<true;
    QTest::newRow("")<<"  A  B "<<true;
}

void XmlspaceTest::testCase1()
{
    QFETCH(int, flag);

    QStringList list;
    list<<""<<" "<<"A"<<"A B"<<" A"<<"B\t"<<"   "<<" A B ";

    QBENCHMARK {
        foreach(QString s, list)
            startsWithOrEndsWithSpace(s, flag);
    }
}

void XmlspaceTest::testCase1_data()
{
    QTest::addColumn<int>("flag");
    QTest::newRow("0") << 0;
    QTest::newRow("1") << 1;
    QTest::newRow("2") << 2;
    QTest::newRow("3") << 3;
    QTest::newRow("4") << 4;
}

QTEST_APPLESS_MAIN(XmlspaceTest)

#include "tst_xmlspacetest.moc"
