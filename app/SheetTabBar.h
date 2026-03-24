#pragma once
// ═══════════════════════════════════════════════════════════════════════════════
//  SheetTabBar.h — Excel-style sheet tabs at the bottom
// ═══════════════════════════════════════════════════════════════════════════════
#include <QWidget>
#include <QList>
#include <QString>

class QScrollBar;
class QToolButton;

class SheetTabBar : public QWidget
{
    Q_OBJECT
public:
    explicit SheetTabBar(QWidget* parent = nullptr);
    ~SheetTabBar() override = default;

    // Tab management
    int  addTab(const QString& name);
    void removeTab(int index);
    void renameTab(int index, const QString& name);
    void setCurrentTab(int index);
    int  currentTab() const { return m_currentIndex; }
    int  tabCount() const   { return m_tabs.size(); }
    QString tabName(int index) const;

    QSize sizeHint() const override;
    QSize minimumSizeHint() const override;

signals:
    void tabActivated(int index);
    void addSheetRequested();
    void renameSheetRequested(int index);
    void deleteSheetRequested(int index);
    void tabMoved(int from, int to);

protected:
    void paintEvent(QPaintEvent* e) override;
    void mousePressEvent(QMouseEvent* e) override;
    void mouseDoubleClickEvent(QMouseEvent* e) override;
    void contextMenuEvent(QContextMenuEvent* e) override;
    void resizeEvent(QResizeEvent* e) override;

private:
    struct TabInfo {
        QString name;
        QRect   rect;
    };

    void relayout();
    int  tabAt(const QPoint& pt) const;

    QList<TabInfo>  m_tabs;
    int             m_currentIndex { 0 };
    int             m_scrollOffset { 0 };
    QToolButton*    m_addButton    { nullptr };
    QToolButton*    m_scrollLeft   { nullptr };
    QToolButton*    m_scrollRight  { nullptr };
};
