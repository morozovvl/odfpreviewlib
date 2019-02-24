#ifndef OdfPreviewLib_H
#define OdfPreviewLib_H

#include <QtCore/QObject>
#include <QtGui/QPainter>
#include <QtPrintSupport/QPrinter>
#include <QtPrintSupport/QPrintPreviewDialog>
#include <QtXml/QDomDocument>

#include "odfpreviewlib_global.h"


enum DocType {none, ods, odt};


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

    void                    drawOds(QPainter*);
    void                    drawOdt(QPainter*);
};

#endif // OdfPreviewLib_H
