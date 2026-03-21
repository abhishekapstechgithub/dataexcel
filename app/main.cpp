#include <QApplication>
#include <QStyleFactory>
#include <QFont>
#include <QPalette>
#include <QDir>
#include <QStandardPaths>
#include "MainWindow.h"

int main(int argc, char* argv[]) {
    // High DPI support (before QApplication construction)
    QApplication::setHighDpiScaleFactorRoundingPolicy(
        Qt::HighDpiScaleFactorRoundingPolicy::PassThrough);

    QApplication app(argc, argv);
    app.setApplicationName("OpenSheet");
    app.setApplicationDisplayName("OpenSheet");
    app.setOrganizationName("OpenSheet");
    app.setApplicationVersion("1.0.0");

    // Fusion style — consistent cross-platform, best base for custom styling
    app.setStyle(QStyleFactory::create("Fusion"));

    // Clean light palette (like WPS)
    QPalette pal;
    pal.setColor(QPalette::Window,          QColor("#f5f5f5"));
    pal.setColor(QPalette::WindowText,      QColor("#222222"));
    pal.setColor(QPalette::Base,            QColor("#ffffff"));
    pal.setColor(QPalette::AlternateBase,   QColor("#f9f9f9"));
    pal.setColor(QPalette::Text,            QColor("#222222"));
    pal.setColor(QPalette::Button,          QColor("#f0f0f0"));
    pal.setColor(QPalette::ButtonText,      QColor("#333333"));
    pal.setColor(QPalette::Highlight,       QColor("#b8d9e8"));
    pal.setColor(QPalette::HighlightedText, QColor("#000000"));
    pal.setColor(QPalette::ToolTipBase,     QColor("#fffbe6"));
    pal.setColor(QPalette::ToolTipText,     QColor("#333333"));
    app.setPalette(pal);

    // Font — Segoe UI like WPS uses on Windows
    QFont f;
    f.setFamily("Segoe UI");
    if (!f.exactMatch()) f.setFamily("Arial");
    f.setPointSize(10);
    app.setFont(f);

    // Tooltip style
    app.setStyleSheet(
        "QToolTip { background:#fffbe6; color:#333; border:1px solid #c8c8c8; "
        "padding:3px 6px; font-size:11px; border-radius:3px; }"
        "QScrollBar:vertical { background:#f5f5f5; width:12px; border:none; }"
        "QScrollBar::handle:vertical { background:#c0c0c0; border-radius:6px; min-height:30px; margin:2px; }"
        "QScrollBar::handle:vertical:hover { background:#a0a0a0; }"
        "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height:0; }"
        "QScrollBar:horizontal { background:#f5f5f5; height:12px; border:none; }"
        "QScrollBar::handle:horizontal { background:#c0c0c0; border-radius:6px; min-width:30px; margin:2px; }"
        "QScrollBar::handle:horizontal:hover { background:#a0a0a0; }"
        "QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width:0; }"
    );

    MainWindow w;
    w.show();
    return app.exec();
}
