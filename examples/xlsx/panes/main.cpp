#include "xlsxdocument.h"
#include "xlsxworksheet.h"

int main()
{
    QXlsx::Document xlsx;

    xlsx.write("A1", "PLANET");
    xlsx.write("A2", "Mercury");
    xlsx.write("A3", "Venus");
    xlsx.write("A4", "Earth");
    xlsx.write("A5", "Mars");
    xlsx.write("A6", "Jupiter");
    xlsx.write("A7", "Saturn");
    xlsx.write("A8", "Uranus");
    xlsx.write("A9", "Neptune");

    xlsx.write("B1", "TYPE");
    xlsx.write("B2", "terrestrial");
    xlsx.write("B3", "terrestrial");
    xlsx.write("B4", "terrestrial");
    xlsx.write("B5", "terrestrial");
    xlsx.write("B6", "gaz giant");
    xlsx.write("B7", "gaz giant");
    xlsx.write("B8", "ice giant");
    xlsx.write("B9", "ice giant");

    xlsx.write("C1", "RADIUS");
    xlsx.write("C2", 2440);
    xlsx.write("C3", 6052);
    xlsx.write("C4", 6378);
    xlsx.write("C5", 3397);
    xlsx.write("C6", 69911);
    xlsx.write("C7", 58232);
    xlsx.write("C8", 25362);
    xlsx.write("C9", 24622);

    xlsx.write("D1", "MASS");
    xlsx.write("D2", 3.301e23);
    xlsx.write("D3", 4.867e24);
    xlsx.write("D4", 6.046e24);
    xlsx.write("D5", 6.417e23);
    xlsx.write("D6", 1.9e27);
    xlsx.write("D7", 5.68e26);
    xlsx.write("D8", 8.68e25);
    xlsx.write("D9", 1.02e26);

    xlsx.renameSheet("Sheet1", "FrozenPanes");
    xlsx.copySheet("FrozenPanes", "SplitPanes");

    xlsx.selectSheet("SplitPanes");
    xlsx.currentWorksheet()->splitPane(1350, 600, "B2");
    xlsx.currentWorksheet()->setSelection("C4");
    xlsx.currentWorksheet()->setAutoFilter("A1:B1");

    xlsx.selectSheet("FrozenPanes");
    xlsx.currentWorksheet()->freezePane("A1", "B2");
    xlsx.currentWorksheet()->setSelection("A4", "A4:D4");
    xlsx.currentWorksheet()->setAutoFilter("A1:B1");

    xlsx.save();
}
