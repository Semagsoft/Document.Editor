#include "app/Application.h"
#include "app/Settings.h"
#include "mainwindow/MainWindow.h"

int main(int argc, char *argv[])
{
    Application app(argc, argv);

    MainWindow mainWindow;
    mainWindow.show();

    return app.exec();
}
