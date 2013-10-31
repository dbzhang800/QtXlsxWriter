#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"

QTXLSX_USE_NAMESPACE

void writeHorizontalAlignCell(Document &xlsx, const QString &cell, const QString &text, Format::HorizontalAlignment align)
{
   Format *format = xlsx.createFormat();
   format->setHorizontalAlignment(align);
   format->setBorderStyle(Format::BorderThin);
   xlsx.write(cell, text, format);
}

void writeVerticalAlignCell(Document &xlsx, const QString &range, const QString &text, Format::VerticalAlignment align)
{
   Format *format = xlsx.createFormat();
   format->setVerticalAlignment(align);
   format->setBorderStyle(Format::BorderThin);
   xlsx.mergeCells(range);

   CellRange r(range);
   for (int row=r.firstRow(); row<=r.lastRow(); ++row) {
       for (int col=r.firstColumn(); col<=r.lastColumn(); ++col) {
           if (row == r.firstRow() && col == r.firstColumn())
               xlsx.write(row, col, text, format);
           else
               xlsx.write(row, col, QVariant(), format);
       }
   }
}

void writeBorderStyleCell(Document &xlsx, const QString &cell, const QString &text, Format::BorderStyle bs)
{
   Format *format = xlsx.createFormat();
   format->setBorderStyle(bs);
   xlsx.write(cell, text, format);
}

void writeSolidFillCell(Document &xlsx, const QString &cell, const QColor &color)
{
   Format *format = xlsx.createFormat();
   format->setPatternBackgroundColor(color);
   xlsx.write(cell, QVariant(), format);
}

void writePatternFillCell(Document &xlsx, const QString &cell, Format::FillPattern pattern, const QColor &color)
{
   Format *format = xlsx.createFormat();
   format->setPatternForegroundColor(color);
   format->setFillPattern(pattern);
   xlsx.write(cell, QVariant(), format);
}

void writeBorderAndFontColorCell(Document &xlsx, const QString &cell, const QString &text, const QColor &color)
{
   Format *format = xlsx.createFormat();
   format->setBorderStyle(Format::BorderThin);
   format->setBorderColor(color);
   format->setFontColor(color);
   xlsx.write(cell, text, format);
}

int main()
{
    Document xlsx;

    //The default sheet is "Sheet1"
    xlsx.setSheetName("Aligns & Borders");
    xlsx.setColumn("B", "B", 20);
    xlsx.setColumn("H", "H", 12);
    xlsx.currentWorksheet()->setGridLinesVisible(false);

    //Alignment
    writeHorizontalAlignCell(xlsx, "B3", "AlignLeft", Format::AlignLeft);
    writeHorizontalAlignCell(xlsx, "B5", "AlignHCenter", Format::AlignHCenter);
    writeHorizontalAlignCell(xlsx, "B7", "AlignRight", Format::AlignRight);
    writeVerticalAlignCell(xlsx, "D3:D7", "AlignTop", Format::AlignTop);
    writeVerticalAlignCell(xlsx, "F3:F7", "AlignVCenter", Format::AlignVCenter);
    writeVerticalAlignCell(xlsx, "H3:H7", "AlignBottom", Format::AlignBottom);

    //Border
    writeBorderStyleCell(xlsx, "B13", "BorderMedium", Format::BorderMedium);
    writeBorderStyleCell(xlsx, "B15", "BorderDashed", Format::BorderDashed);
    writeBorderStyleCell(xlsx, "B17", "BorderDotted", Format::BorderDotted);
    writeBorderStyleCell(xlsx, "B19", "BorderThick", Format::BorderThick);
    writeBorderStyleCell(xlsx, "B21", "BorderDouble", Format::BorderDouble);
    writeBorderStyleCell(xlsx, "B23", "BorderDashDot", Format::BorderDashDot);

    //Fill
    writeSolidFillCell(xlsx, "D13", Qt::red);
    writeSolidFillCell(xlsx, "D15", Qt::blue);
    writeSolidFillCell(xlsx, "D17", Qt::yellow);
    writeSolidFillCell(xlsx, "D19", Qt::magenta);
    writeSolidFillCell(xlsx, "D21", Qt::green);
    writeSolidFillCell(xlsx, "D23", Qt::gray);
    writePatternFillCell(xlsx, "F13", Format::PatternMediumGray, Qt::red);
    writePatternFillCell(xlsx, "F15", Format::PatternDarkHorizontal, Qt::blue);
    writePatternFillCell(xlsx, "F17", Format::PatternDarkVertical, Qt::yellow);
    writePatternFillCell(xlsx, "F19", Format::PatternDarkDown, Qt::magenta);
    writePatternFillCell(xlsx, "F21", Format::PatternLightVertical, Qt::green);
    writePatternFillCell(xlsx, "F23", Format::PatternLightTrellis, Qt::gray);

    writeBorderAndFontColorCell(xlsx, "H13", "Qt::red", Qt::red);
    writeBorderAndFontColorCell(xlsx, "H15", "Qt::blue", Qt::blue);
    writeBorderAndFontColorCell(xlsx, "H17", "Qt::yellow", Qt::yellow);
    writeBorderAndFontColorCell(xlsx, "H19", "Qt::magenta", Qt::magenta);
    writeBorderAndFontColorCell(xlsx, "H21", "Qt::green", Qt::green);
    writeBorderAndFontColorCell(xlsx, "H23", "Qt::gray", Qt::gray);

    xlsx.saveAs("Book1.xlsx");

    //Make sure that read/write works well.
    Document xlsx2("Book1.xlsx");
    xlsx2.saveAs("Book2.xlsx");

    return 0;
}
