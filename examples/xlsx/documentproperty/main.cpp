#include <QtCore>
#include "xlsxworkbook.h"

#ifdef Q_OS_MAC
#  define DATA_PATH "../../../"
#else
#  define DATA_PATH "./"
#endif

int main()
{
    QXlsx::Workbook workbook;
/*
    These properties are visible when you use the
    Office Button -> Prepare -> Properties option in Excel and are also
    available to external applications that read or index windows files
*/
    workbook.setProperty("title", "This is an example spreadsheet");
    workbook.setProperty("subject", "With document properties");
    workbook.setProperty("creator", "Debao Zhang");
    workbook.setProperty("company", "HMICN");
    workbook.setProperty("category", "Example spreadsheets");
    workbook.setProperty("keywords", "Sample, Example, Properties");
    workbook.setProperty("description", "Created with Qt Xlsx");

    workbook.save(DATA_PATH"Test.xlsx");
    return 0;
}
