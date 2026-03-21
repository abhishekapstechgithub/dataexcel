#include "ColorButton.h"
#include <QPainter>
#include <QColorDialog>

ColorButton::ColorButton(const QString& label, QWidget* parent)
    : QPushButton(label, parent), m_label(label), m_color(Qt::black)
{
    setFixedHeight(26);
    connect(this, &QPushButton::clicked, this, [this]() {
        QColor c = QColorDialog::getColor(m_color, this, "Choose Color");
        if (c.isValid()) { setColor(c); emit colorChanged(c); }
    });
}

void ColorButton::setColor(const QColor& c) {
    m_color = c;
    update();
}

void ColorButton::paintEvent(QPaintEvent* e) {
    QPushButton::paintEvent(e);
    QPainter p(this);
    // Draw colored strip at the bottom
    QRect strip(4, height()-5, width()-8, 4);
    p.fillRect(strip, m_color);
}
