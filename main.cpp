#include <QDebug>
#include <QtWidgets/QApplication>

#include "odfpreviewlib.h"


int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    OdfPreviewLib ods;
    if (ods.open("./nakl.ods"))
    {
        ods.preview();
        ods.close();
    }
}
