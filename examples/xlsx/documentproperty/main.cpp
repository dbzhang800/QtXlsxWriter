#include <QtCore>
#include "xlsxdocument.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main()
{
    QXlsx::Document xlsx;
/*
    These properties are visible when you use the
    Office Button -> Prepare -> Properties option in Excel and are also
    available to external applications that read or index windows files
*/
    xlsx.setDocumentProperty("title", "This is an example spreadsheet");
    xlsx.setDocumentProperty("subject", "With document properties");
    xlsx.setDocumentProperty("creator", "Debao Zhang");
    xlsx.setDocumentProperty("company", "HMICN");
    xlsx.setDocumentProperty("category", "Example spreadsheets");
    xlsx.setDocumentProperty("keywords", "Sample, Example, Properties");
    xlsx.setDocumentProperty("description", "Created with Qt Xlsx");

    xlsx.saveAs(DATA_PATH"Test.xlsx");
    return 0;
}
