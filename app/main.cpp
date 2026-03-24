// ═══════════════════════════════════════════════════════════════════════════════
//  main.cpp — OpenSheet application entry point
// ═══════════════════════════════════════════════════════════════════════════════
#include <QApplication>
#include <QSurfaceFormat>
#include <QStyleFactory>
#include <QFontDatabase>
#include "MainWindow.h"

int main(int argc, char* argv[])
{
    // High-DPI support
    QApplication::setHighDpiScaleFactorRoundingPolicy(
        Qt::HighDpiScaleFactorRoundingPolicy::PassThrough);

    QApplication app(argc, argv);
    app.setApplicationName("OpenSheet");
    app.setApplicationDisplayName("OpenSheet");
    app.setApplicationVersion("1.0.0");
    app.setOrganizationName("OpenSheet");
    app.setOrganizationDomain("opensheet.app");

    // Use Fusion style for cross-platform consistent look
    app.setStyle(QStyleFactory::create("Fusion"));

    // Global palette (light theme matching Excel)
    QPalette palette;
    palette.setColor(QPalette::Window,          QColor(0xF2F2F2));
    palette.setColor(QPalette::WindowText,      QColor(0x212121));
    palette.setColor(QPalette::Base,            Qt::white);
    palette.setColor(QPalette::AlternateBase,   QColor(0xF8F8F8));
    palette.setColor(QPalette::Text,            QColor(0x212121));
    palette.setColor(QPalette::Button,          QColor(0xF2F2F2));
    palette.setColor(QPalette::ButtonText,      QColor(0x212121));
    palette.setColor(QPalette::Highlight,       QColor(0x217346));
    palette.setColor(QPalette::HighlightedText, Qt::white);
    palette.setColor(QPalette::Link,            QColor(0x217346));
    app.setPalette(palette);

    // Default application font
    QFont appFont("Calibri", 10);
    if (!QFontDatabase::families().contains("Calibri"))
        appFont = QFont("Segoe UI", 10);
    app.setFont(appFont);

    // Open file passed on command line (if any)
    QString fileToOpen;
    if (argc > 1)
        fileToOpen = QString::fromLocal8Bit(argv[1]);

    MainWindow w;
    w.show();

    if (!fileToOpen.isEmpty())
        w.openFile(fileToOpen);

    return app.exec();
}
