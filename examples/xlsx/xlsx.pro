TEMPLATE = subdirs
SUBDIRS = hello \
    extractdata \
    style \
    documentproperty \
    image \
    mergecells \
    rowcolumn \
    numberformat \
    datavalidation \
    definename \
    formulas \
    richtext \
    conditionalformatting \
    worksheetoperations \
    hyperlinks \
    chart \
    chartsheet \
    calendar \
    demo \
    histogram

qtHaveModule(widgets): SUBDIRS += xlsxwidget

