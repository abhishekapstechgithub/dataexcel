#pragma once
#include <QPushButton>
#include <QColor>

// A button that shows a colored swatch and opens a QColorDialog on click
class ColorButton : public QPushButton {
    Q_OBJECT
    Q_PROPERTY(QColor color READ color WRITE setColor NOTIFY colorChanged)
public:
    explicit ColorButton(const QString& label, QWidget* parent = nullptr);
    QColor color() const { return m_color; }
    void   setColor(const QColor& c);
signals:
    void colorChanged(const QColor& c);
protected:
    void paintEvent(QPaintEvent* e) override;
private:
    QColor  m_color;
    QString m_label;
};
