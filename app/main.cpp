#include <QApplication>
#include <QStyleFactory>
#include <QFont>
#include <QPalette>
#include "MainWindow.h"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);

    // Application metadata
    app.setApplicationName("OpenSheet");
    app.setOrganizationName("OpenSheet");
    app.setApplicationVersion("1.0.0");

    // Use Fusion style for consistent cross-platform look
    app.setStyle(QStyleFactory::create("Fusion"));

    // Set default font (Calibri-like)
    QFont defaultFont("Segoe UI", 10);
    if (!defaultFont.exactMatch()) defaultFont.setFamily("Arial");
    app.setFont(defaultFont);

    // Light palette base
    QPalette pal = app.palette();
    pal.setColor(QPalette::Window,          QColor("#f5f5f5"));
    pal.setColor(QPalette::WindowText,      QColor("#222222"));
    pal.setColor(QPalette::Base,            QColor("#ffffff"));
    pal.setColor(QPalette::AlternateBase,   QColor("#f9f9f9"));
    pal.setColor(QPalette::Highlight,       QColor("#c0d8f0"));
    pal.setColor(QPalette::HighlightedText, QColor("#000000"));
    app.setPalette(pal);

    MainWindow w;
    w.show();
    return app.exec();
}
