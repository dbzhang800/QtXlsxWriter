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

#include "xlsxconditionalformatting.h"
#include "xlsxconditionalformatting_p.h"
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"

#include <QXmlStreamReader>
#include <QXmlStreamWriter>
#include <QDebug>

QT_BEGIN_NAMESPACE_XLSX

ConditionalFormattingPrivate::ConditionalFormattingPrivate()
{

}

ConditionalFormattingPrivate::ConditionalFormattingPrivate(const ConditionalFormattingPrivate &other)
    :QSharedData(other)
{

}

ConditionalFormattingPrivate::~ConditionalFormattingPrivate()
{

}

/*!
 * \class ConditionalFormatting
 * \brief Conditional formatting for single cell or ranges
 * \inmodule QtXlsx
 *
 * The conditional formatting can be applied to a single cell or ranges of cells.
 */


/*!
    \enum ConditionalFormatting::HighlightRuleType

    \value Highlight_LessThan
    \value Highlight_LessThanOrEqual
    \value Highlight_Equal
    \value Highlight_NotEqual
    \value Highlight_GreaterThanOrEqual
    \value Highlight_GreaterThan
    \value Highlight_Between
    \value Highlight_NotBetween

    \value Highlight_ContainsText
    \value Highlight_NotContainsText
    \value Highlight_BeginsWith
    \value Highlight_EndsWith

    \value Highlight_TimePeriod

    \value Highlight_Duplicate
    \value Highlight_Unique

    \value Highlight_Blanks
    \value Highlight_NoBlanks
    \value Highlight_Errors
    \value Highlight_NoErrors

    \value Highlight_Top
    \value Highlight_TopPercent
    \value Highlight_Bottom
    \value Highlight_BottomPercent

    \value Highlight_AboveAverage
    \value Highlight_AboveOrEqualAverage
    \value Highlight_BelowAverage
    \value Highlight_BelowOrEqualAverage
    \value Highlight_AboveStdDev1
    \value Highlight_AboveStdDev2
    \value Highlight_AboveStdDev3
    \value Highlight_BelowStdDev1
    \value Highlight_BelowStdDev2
    \value Highlight_BelowStdDev3

    \value Highlight_SatisfyFormula
*/

/*!
    Construct a conditional formatting object
*/
ConditionalFormatting::ConditionalFormatting()
    :d(new ConditionalFormattingPrivate())
{

}

/*!
    Constructs a copy of \a other.
*/
ConditionalFormatting::ConditionalFormatting(const ConditionalFormatting &other)
    :d(other.d)
{

}

/*!
    Assigns \a other to this conditional formatting and returns a reference to
    this conditional formatting.
 */
ConditionalFormatting &ConditionalFormatting::operator=(const ConditionalFormatting &other)
{
    this->d = other.d;
    return *this;
}


/*!
 * Destroy the object.
 */
ConditionalFormatting::~ConditionalFormatting()
{
}

/*!
 * Add a hightlight rule with the given \a type, \a formula1, \a formula2,
 * \a format and \a stopIfTrue.
 */
bool ConditionalFormatting::addHighlightCellsRule(HighlightRuleType type, const QString &formula1, const QString &formula2, const Format &format, bool stopIfTrue)
{
    if (format.isEmpty())
        return false;

    bool skipFormula1 = false;

    QSharedPointer<XlsxCfRuleData> cfRule(new XlsxCfRuleData);
    if (type >= Highlight_LessThan && type <= Highlight_NotBetween) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("cellIs");
        QString op;
        switch (type) {
        case Highlight_Between: op = QStringLiteral("between"); break;
        case Highlight_Equal: op = QStringLiteral("equal"); break;
        case Highlight_GreaterThan: op = QStringLiteral("greaterThan"); break;
        case Highlight_GreaterThanOrEqual: op = QStringLiteral("greaterThanOrEqual"); break;
        case Highlight_LessThan: op = QStringLiteral("lessThan"); break;
        case Highlight_LessThanOrEqual: op = QStringLiteral("lessThanOrEqual"); break;
        case Highlight_NotBetween: op = QStringLiteral("notBetween"); break;
        case Highlight_NotEqual: op = QStringLiteral("notEqual"); break;
        default: break;
        }
        cfRule->attrs[XlsxCfRuleData::A_operator] = op;
    } else if (type >= Highlight_ContainsText && type <= Highlight_EndsWith) {
        if (type == Highlight_ContainsText) {
            cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("containsText");
            cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("containsText");
            cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("NOT(ISERROR(SEARCH(\"%1\",%2)))").arg(formula1);
        } else if (type == Highlight_NotContainsText) {
            cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("notContainsText");
            cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("notContains");
            cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("ISERROR(SEARCH(\"%2\",%1))").arg(formula1);
        } else if (type == Highlight_BeginsWith) {
            cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("beginsWith");
            cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("beginsWith");
            cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("LEFT(%2,LEN(\"%1\"))=\"%1\"").arg(formula1);
        } else {
            cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("endsWith");
            cfRule->attrs[XlsxCfRuleData::A_operator] = QStringLiteral("endsWith");
            cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("RIGHT(%2,LEN(\"%1\"))=\"%1\"").arg(formula1);
        }
        cfRule->attrs[XlsxCfRuleData::A_text] = formula1;
        skipFormula1 = true;
    } else if (type == Highlight_TimePeriod) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("timePeriod");
        //:Todo
        return false;
    } else if (type == Highlight_Duplicate) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("duplicateValues");
    } else if (type == Highlight_Unique) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("uniqueValues");
    } else if (type == Highlight_Errors) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("containsErrors");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("ISERROR(%1)");
        skipFormula1 = true;
    } else if (type == Highlight_NoErrors) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("notContainsErrors");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("NOT(ISERROR(%1))");
        skipFormula1 = true;
    } else if (type == Highlight_Blanks) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("containsBlanks");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("LEN(TRIM(%1))=0");
        skipFormula1 = true;
    } else if (type == Highlight_NoBlanks) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("notContainsBlanks");
        cfRule->attrs[XlsxCfRuleData::A_formula1_temp] = QStringLiteral("LEN(TRIM(%1))>0");
        skipFormula1 = true;
    } else if (type >= Highlight_Top && type <= Highlight_BottomPercent) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("top10");
        if (type == Highlight_Bottom || type == Highlight_BottomPercent)
            cfRule->attrs[XlsxCfRuleData::A_bottom] = QStringLiteral("1");
        if (type == Highlight_TopPercent || type == Highlight_BottomPercent)
            cfRule->attrs[XlsxCfRuleData::A_percent] = QStringLiteral("1");
        cfRule->attrs[XlsxCfRuleData::A_rank] = !formula1.isEmpty() ? formula1 : QStringLiteral("10");
        skipFormula1 = true;
    } else if (type >= Highlight_AboveAverage && type <= Highlight_BelowStdDev3) {
        cfRule->attrs[XlsxCfRuleData::A_type] = QStringLiteral("aboveAverage");
        if (type >= Highlight_BelowAverage && type <= Highlight_BelowStdDev3)
            cfRule->attrs[XlsxCfRuleData::A_aboveAverage] = QStringLiteral("0");
        if (type == Highlight_AboveOrEqualAverage || type == Highlight_BelowOrEqualAverage)
            cfRule->attrs[XlsxCfRuleData::A_equalAverage] = QStringLiteral("1");
        if (type == Highlight_AboveStdDev1 || type == Highlight_BelowStdDev1)
            cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("1");
        else if (type == Highlight_AboveStdDev2 || type == Highlight_BelowStdDev2)
            cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("2");
        else if (type == Highlight_AboveStdDev3 || type == Highlight_BelowStdDev3)
            cfRule->attrs[XlsxCfRuleData::A_stdDev] = QStringLiteral("3");
    } else {
        return false;
    }

    cfRule->dxfFormat = format;
    if (stopIfTrue)
        cfRule->attrs[XlsxCfRuleData::A_stopIfTrue] = true;
    if (!formula1.isEmpty() && !skipFormula1)
        cfRule->attrs[XlsxCfRuleData::A_formula1] = formula1.startsWith(QLatin1String("=")) ? formula1.mid(1) : formula1;
    if (!formula2.isEmpty())
        cfRule->attrs[XlsxCfRuleData::A_formula2] = formula2.startsWith(QLatin1String("=")) ? formula2.mid(1) : formula2;

    d->cfRules.append(cfRule);
    return true;
}

/*!
 * \overload
 *
 * Add a hightlight rule with the given \a type \a format and \a stopIfTrue.
 */
bool ConditionalFormatting::addHighlightCellsRule(HighlightRuleType type, const Format &format, bool stopIfTrue)
{
    if ((type >= Highlight_AboveAverage && type <= Highlight_BelowStdDev3)
            || (type >= Highlight_Duplicate && type <= Highlight_NoErrors)) {
        return addHighlightCellsRule(type, QString(), QString(), format, stopIfTrue);
    }

    return false;
}

/*!
 * \overload
 *
 * Add a hightlight rule with the given \a type, \a formula, \a format and \a stopIfTrue.
 */
bool ConditionalFormatting::addHighlightCellsRule(HighlightRuleType type, const QString &formula, const Format &format, bool stopIfTrue)
{
    if (type == Highlight_Between || type == Highlight_NotBetween)
        return false;

    return addHighlightCellsRule(type, formula, QString(), format, stopIfTrue);
}

/*!
    Returns the ranges on which the validation will be applied.
 */
QList<CellRange> ConditionalFormatting::ranges() const
{
    return d->ranges;
}

/*!
    Add the \a cell on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addCell(const QString &cell)
{
    d->ranges.append(CellRange(cell));
}

/*!
    \overload
    Add the cell(\a row, \a col) on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addCell(int row, int col)
{
    d->ranges.append(CellRange(row, col, row, col));
}

/*!
    Add the \a range on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addRange(const QString &range)
{
    d->ranges.append(CellRange(range));
}

/*!
    \overload
    Add the range(\a firstRow, \a firstCol, \a lastRow, \a lastCol) on
    which the conditional formatting will apply to.
 */
void ConditionalFormatting::addRange(int firstRow, int firstCol, int lastRow, int lastCol)
{
    d->ranges.append(CellRange(firstRow, firstCol, lastRow, lastCol));
}

/*!
    \overload
    Add the \a range on which the conditional formatting will apply to.
 */
void ConditionalFormatting::addRange(const CellRange &range)
{
    d->ranges.append(range);
}

bool ConditionalFormatting::loadFromXml(QXmlStreamReader &reader) const
{
    Q_ASSERT(reader.name() == QStringLiteral("conditionalFormatting"));

    return false;
}

bool ConditionalFormatting::saveToXml(QXmlStreamWriter &writer) const
{
    writer.writeStartElement(QStringLiteral("conditionalFormatting"));
    QStringList sqref;
    foreach (CellRange range, ranges())
        sqref.append(range.toString());
    writer.writeAttribute(QStringLiteral("sqref"), sqref.join(QLatin1Char(' ')));

    for (int i=0; i<d->cfRules.size(); ++i) {
        const QSharedPointer<XlsxCfRuleData> &rule = d->cfRules[i];
        writer.writeStartElement(QStringLiteral("cfRule"));
        writer.writeAttribute(QStringLiteral("type"), rule->attrs[XlsxCfRuleData::A_type].toString());
        if (rule->dxfFormat.dxfIndexValid())
            writer.writeAttribute(QStringLiteral("dxfId"), QString::number(rule->dxfFormat.dxfIndex()));
        writer.writeAttribute(QStringLiteral("priority"), QString::number(rule->priority));
        if (rule->attrs.contains(XlsxCfRuleData::A_stopIfTrue))
            writer.writeAttribute(QStringLiteral("stopIfTrue"), rule->attrs[XlsxCfRuleData::A_stopIfTrue].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_aboveAverage))
            writer.writeAttribute(QStringLiteral("aboveAverage"), rule->attrs[XlsxCfRuleData::A_aboveAverage].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_percent))
            writer.writeAttribute(QStringLiteral("percent"), rule->attrs[XlsxCfRuleData::A_percent].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_bottom))
            writer.writeAttribute(QStringLiteral("bottom"), rule->attrs[XlsxCfRuleData::A_bottom].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_operator))
            writer.writeAttribute(QStringLiteral("operator"), rule->attrs[XlsxCfRuleData::A_operator].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_text))
            writer.writeAttribute(QStringLiteral("text"), rule->attrs[XlsxCfRuleData::A_text].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_timePeriod))
            writer.writeAttribute(QStringLiteral("timePeriod"), rule->attrs[XlsxCfRuleData::A_timePeriod].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_rank))
            writer.writeAttribute(QStringLiteral("rank"), rule->attrs[XlsxCfRuleData::A_rank].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_stdDev))
            writer.writeAttribute(QStringLiteral("stdDev"), rule->attrs[XlsxCfRuleData::A_stdDev].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_equalAverage))
            writer.writeAttribute(QStringLiteral("equalAverage"), rule->attrs[XlsxCfRuleData::A_equalAverage].toString());

        if (rule->attrs.contains(XlsxCfRuleData::A_formula1_temp)) {
            QString startCell = ranges()[0].toString().split(QLatin1Char(':'))[0];
            writer.writeTextElement(QStringLiteral("formula"), rule->attrs[XlsxCfRuleData::A_formula1_temp].toString().arg(startCell));
        } else if (rule->attrs.contains(XlsxCfRuleData::A_formula1)) {
            writer.writeTextElement(QStringLiteral("formula"), rule->attrs[XlsxCfRuleData::A_formula1].toString());
        }
        if (rule->attrs.contains(XlsxCfRuleData::A_formula2))
            writer.writeTextElement(QStringLiteral("formula"), rule->attrs[XlsxCfRuleData::A_formula2].toString());
        if (rule->attrs.contains(XlsxCfRuleData::A_formula3))
            writer.writeTextElement(QStringLiteral("formula"), rule->attrs[XlsxCfRuleData::A_formula3].toString());

        writer.writeEndElement(); //cfRule
    }

    writer.writeEndElement(); //conditionalFormatting
    return true;
}

QT_END_NAMESPACE_XLSX
