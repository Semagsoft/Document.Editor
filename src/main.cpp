#include "app/Application.h"
#include "app/Settings.h"
#include "mainwindow/MainWindow.h"

int main(int argc, char *argv[])
{
    Application app(argc, argv);

    MainWindow mainWindow;
    mainWindow.show();

    for (const QString& file : app.startupFiles())
        mainWindow.openFile(file);

    return app.exec();
}
