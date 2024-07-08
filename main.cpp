#include "mainwindow.h"
#include <QApplication>
#include<QtWidgets>
int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    a.setOrganizationName("Excel_");
    a.setApplicationName("_excelSettings");
    QSplashScreen *screen=new QSplashScreen(QPixmap(":/icon/icon/iconExcel.png"));
    screen->show();
    MainWindow w;
    screen->close();
    w.show();
    return a.exec();
}
