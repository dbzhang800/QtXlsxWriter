#ifndef QXLSXWRITER_H
#define QXLSXWRITER_H

#include <QObject>

class QXlsxWriter : public QObject
{
    Q_OBJECT
public:
    explicit QXlsxWriter(QObject *parent = 0);

//    void worksheets();
    
signals:
    
public slots:
//    void addWorksheet(QString name);
//    void addFormat();
//    void addChart();
//    void setProperties();
};

#endif // QXLSXWRITER_H
