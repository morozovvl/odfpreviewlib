#include <QDebug>
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
}


void OdfPreviewLib::drawOds(QPainter* painter)
{
    qDebug() << painter->device()->height() << painter->device()->width();
/*
    QHash<int, int> columnWidths;
    QHash<int, int> rowHeights;

    QDomNodeList styles = document.elementsByTagName("office:automatic-styles");
    QDomNodeList nodeList = document.elementsByTagName("table:table-row");

    int vPos = 0;       // vertical position by mm's;
    int hPos = 0;       // horizontal position by mm's;

    for (int c = 0; document.elementsByTagName("table:table-column").count(); c++)
    {
        columnWidths.insert(c, )
    }

    for (int i = 0; i < nodeList.count(); i++)
    {
        painter->setFont(QFont("Tahoma",8));
        painter->drawText(100, 100, "It's the spreadsheet!");
    }
*/
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
