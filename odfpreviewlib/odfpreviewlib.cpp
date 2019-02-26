#include "odfpreviewlib.h"
#include "quazip/quazip.h"
#include "quazip/quazipfile.h"


OdfPreviewLib::OdfPreviewLib(QWidget *parent) : QObject()
{
    printer         = new QPrinter();
    printPreview    = new QPrintPreviewDialog(printer, parent);

    printPreview->setWindowTitle("Preview Dialog");
    Qt::WindowFlags flags(Qt::WindowTitleHint);
    printPreview->setWindowFlags(flags);

    connect(printPreview, SIGNAL(paintRequested(QPrinter*)), this, SLOT(draw(QPrinter*)));
}


OdfPreviewLib::~OdfPreviewLib()
{
    disconnect(printPreview, SIGNAL(paintRequested(QPrinter*)), this, SLOT(draw(QPrinter*)));

    delete printPreview;
    delete printer;
}


bool OdfPreviewLib::open(QString fileName)
{
    bool lResult = false;

    if (printer->isValid())
    {
        if (unzip(fileName))
        {
            lResult = true;
        }
    }

    return lResult;
}


bool OdfPreviewLib::open(QDomDocument* doc)
{
    document = *doc;
    return true;
}


DocType OdfPreviewLib::getDocType()
{
    DocType result = none;

    if (document.elementsByTagName("office:spreadsheet").count() > 0)
        result = ods;
    else if (document.elementsByTagName("office:text").count() > 0)
        result = odt;

    return result;
}


void OdfPreviewLib::close()
{
}


void OdfPreviewLib::preview()
{
    setPrinterConfig();
    printPreview->exec();
}


void OdfPreviewLib::print()
{
    setPrinterConfig();
}


void OdfPreviewLib::draw(QPrinter *printer)
{
    QPainter painter(printer);
    painter.setRenderHints(QPainter::Antialiasing |
                       QPainter::TextAntialiasing |
                       QPainter::SmoothPixmapTransform, true);

    switch (getDocType())
    {
        case ods:
            drawOds(&painter);
            break;
        case odt:
            drawOdt(&painter);
            break;
        default:
            break;
    }
}


bool OdfPreviewLib::unzip(QString fileName)
{
    bool lResult = false;
    QuaZip zip(fileName);
    zip.open(QuaZip::mdUnzip);

    if (zip.setCurrentFile("content.xml"))
    {
        QuaZipFile file(&zip);
        file.open(QIODevice::ReadOnly);

        QByteArray data(file.size(), ' ');
        file.read(data.data(), file.size());

        if (document.setContent(data))
            lResult = true;

        file.close();
    }
    zip.close();
    return lResult;
}


void OdfPreviewLib::setPrinterConfig()
{
    printer->setPageSize(QPrinter::A4);
    printer->setOrientation(QPrinter::Portrait);
    printer->setFullPage(true);
    printer->setPageMargins(10.0, 10.0, 10.0, 10.0, QPrinter::Millimeter);
}


void OdfPreviewLib::drawOds(QPainter* painter)
{
    // Loading styles of rows, columns, cells from document
    QDomNodeList    nl = document.elementsByTagName("style:style");
    for (int i = 0; i < nl.count(); i++)
    {
        QString name = nl.at(i).toElement().attribute("style:name");
        QString styleFamily = nl.at(i).toElement().attribute("style:family");
        QString alignment = nl.at(i).firstChildElement("style:paragraph-properties").attribute("fo:text-align");
        Style style;

        style.width = 0;
        style.height = 0;
        style.fontName = "";
        style.fontSize = 0;
        style.align = Qt::AlignLeft;

        if (styleFamily == "table")
        {
            style.type = tableTable;
        }
        else if (styleFamily == "table-row")
        {
            style.type = tableRow;
            style.height = QString(nl.at(i).firstChild().toElement().attribute("style:row-height").remove("mm")).toFloat();
        }
        else if (styleFamily == "table-column")
        {
            style.type = tableColumn;
            style.width = QString(nl.at(i).firstChild().toElement().attribute("style:column-width").remove("mm")).toFloat();
        }
        else if (styleFamily == "table-cell")
        {
            style.type = tableCell;
            style.fontName = nl.at(i).firstChildElement("style:text-properties").attribute("style:font-name");
            style.fontSize = QString(nl.at(i).firstChildElement("style:text-properties").attribute("fo:font-size").remove("pt")).toInt();
        }
        else
        {
            style.type = tableNone;
        }

        if (alignment == "start")
            style.align = Qt::AlignLeft;
        else if (alignment == "end")
            style.align = Qt::AlignRight;
        if (alignment == "center")
            style.align = Qt::AlignHCenter;

        if (style.type != tableNone)
            styles.insert(name, style);
    }

    // Calculate row's vertical position and height
    QDomNodeList    rows = document.elementsByTagName("table:table-row");
    float   y = 0;
    for (int i = 0; i < rows.count(); i++)
    {
        RowPos pos;

        pos.y = y;
        pos.h = styles[rows.at(i).toElement().attribute("table:style-name")].height;

        rowsPos.insert(i, pos);
        y += pos.h;
    }

    // Calculate column's horizontal position and width
    QDomNodeList    columns = document.elementsByTagName("table:table-column");
    float   x = 0;
    for (int i = 0; i < columns.count(); i++)
    {
        ColumnPos pos;

        pos.x = x;
        pos.w = styles[columns.at(i).toElement().attribute("table:style-name")].width;

        columnsPos.insert(i, pos);
        x += pos.w;
    }

    // Calculate cell's positions
    float rowY = 0;
    float rowH = 0;

    rows = document.elementsByTagName("table:table-row");
    for (int i = 0; i < rows.count(); i++)
    {
        float colX = 0;
        float colW = 0;
        QDomNodeList cells = rows.at(i).childNodes();
        for (int j = 0; j < cells.count(); j++)
        {
            if (cells.at(j).hasChildNodes())
            {
                QString styleName = cells.at(j).toElement().attribute("table:style-name");
                QString text = cells.at(j).firstChildElement("text:p").text();
                int rowSpanned = QString(cells.at(j).toElement().attribute("table:number-rows-spanned")).toInt();
                int colSpanned = QString(cells.at(j).toElement().attribute("table:number-columns-spanned")).toInt();

                rowY = rowsPos[i].y;
                if (rowSpanned > 1)
                {
                    for (int s = 0; s < rowSpanned; s++)
                    {
                        rowH += rowsPos[i+s].h;
                    }
                }
                else
                    rowH = rowsPos[i].h;

                colX = columnsPos[j].x;
                if (colSpanned > 0)
                {
                    for (int s = 0; s < colSpanned; s++)
                    {
                        colW += columnsPos[j+s].w;
                    }
                }
                else
                    colW = columnsPos[j].w;

                // Draw each cell of table

                QRect rect(mmToPixels(colX), mmToPixels(rowY), mmToPixels(colW), mmToPixels(rowH));
                QRect boundingRect;
                painter->drawText(rect, Qt::AlignVCenter | styles[styleName].align, text, &boundingRect);

            }
        }
    }
}


void OdfPreviewLib::drawOdt(QPainter* painter)
{
    painter->setFont(QFont("Tahoma",8));
    painter->drawText(100, 100, "It's the text!");
}


int OdfPreviewLib::strMmToPix(QString str)
{
    str = str.replace("mm", "");
    return 0;
}


qreal OdfPreviewLib::mmToPixels(qreal mm)
{
    return mm * 0.039370147 * printer->resolution();
}

