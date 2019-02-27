#ifndef OdfPreviewLib_H
#define OdfPreviewLib_H

#include <QtCore/QObject>
#include <QtGui/QPainter>
#include <QtPrintSupport/QPrinter>
#include <QtPrintSupport/QPrintPreviewDialog>
#include <QtXml/QDomDocument>

#include "odfpreviewlib_global.h"


enum DocType {none, ods, odt};
enum StyleFamily {tableNone, tableTable, tableRow, tableColumn, tableCell};

struct BorderStyle
{
    float       size;
    QString     type;
    QString     color;
};

struct Style
{
    StyleFamily         type;
    float               width;
    float               height;
    QString             fontName;
    int                 fontSize;
    Qt::AlignmentFlag   align;
    BorderStyle         leftBS;     // left border style
    BorderStyle         rightBS;
    BorderStyle         topBS;
    BorderStyle         bottomBS;
};

struct RowPos
{
    float       y;
    float       h;
};

struct ColumnPos
{
    float       x;
    float       w;
};

class OdfPreviewLibSHARED_EXPORT OdfPreviewLib : public QObject
{
    Q_OBJECT

public:
    OdfPreviewLib(QWidget *parent = nullptr);
    ~OdfPreviewLib();

    bool open(QString);
    bool open(QDomDocument*);
    void close();
    void preview();
    void print();

private slots:
    void draw(QPrinter*);

private:
    QPrinter*               printer;
    QPrintPreviewDialog*    printPreview;
    QDomDocument            document;

    bool                    unzip(QString);
    DocType                 getDocType();
    void                    setPrinterConfig();
    int                     strMmToPix(QString);
    QHash<QString, Style>   styles;
    QHash<int, RowPos>   rowsPos;
    QHash<int, ColumnPos>   columnsPos;

    void                    drawOds(QPainter*);
    void                    drawOdt(QPainter*);

    qreal                   mmToPixels(qreal);
    BorderStyle             parseBorderTypeString(QString);
};

#endif // OdfPreviewLib_H
