#include <QApplication>
#include <QStyleFactory>
#include "MainWindow.h"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);
    app.setApplicationName("QtSpreadsheet");
    app.setOrganizationName("QtSpreadsheet");
    app.setApplicationVersion("1.0.0");

    // Use Fusion style for consistent cross-platform look
    app.setStyle(QStyleFactory::create("Fusion"));

    // Accent color palette (Excel-inspired green)
    QPalette pal = app.palette();
    pal.setColor(QPalette::Highlight,     QColor("#217346"));
    pal.setColor(QPalette::HighlightedText, Qt::white);
    app.setPalette(pal);

    MainWindow w;
    w.show();
    return app.exec();
}
