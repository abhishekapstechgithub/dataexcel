// ═══════════════════════════════════════════════════════════════════════════════
//  SheetTabBar.cpp — Excel-style sheet tabs implementation
// ═══════════════════════════════════════════════════════════════════════════════
#include "SheetTabBar.h"
#include <QPainter>
#include <QPainterPath>
#include <QMouseEvent>
#include <QContextMenuEvent>
#include <QMenu>
#include <QAction>
#include <QInputDialog>
#include <QToolButton>
#include <QHBoxLayout>
#include <QResizeEvent>
#include <QFontMetrics>
#include <algorithm>

static constexpr int TAB_HEIGHT   = 26;
static constexpr int TAB_PADDING  = 14;
static constexpr int MIN_TAB_W    = 60;
static constexpr int MAX_TAB_W    = 160;
static constexpr int BTN_W        = 22;

SheetTabBar::SheetTabBar(QWidget* parent)
    : QWidget(parent)
{
    setFixedHeight(TAB_HEIGHT + 2);
    setStyleSheet("background: #D0D0D0; border-top: 1px solid #B0B0B0;");

    // Scroll left / right buttons (for overflow)
    m_scrollLeft = new QToolButton(this);
    m_scrollLeft->setText("<");
    m_scrollLeft->setFixedSize(BTN_W, TAB_HEIGHT);
    m_scrollLeft->hide();
    connect(m_scrollLeft, &QToolButton::clicked, this, [this]{
        m_scrollOffset = std::max(0, m_scrollOffset - 1);
        relayout(); update();
    });

    m_scrollRight = new QToolButton(this);
    m_scrollRight->setText(">");
    m_scrollRight->setFixedSize(BTN_W, TAB_HEIGHT);
    m_scrollRight->hide();
    connect(m_scrollRight, &QToolButton::clicked, this, [this]{
        m_scrollOffset = std::min((int)m_tabs.size() - 1, m_scrollOffset + 1);
        relayout(); update();
    });

    // "+" button to add sheets
    m_addButton = new QToolButton(this);
    m_addButton->setText("+");
    m_addButton->setFixedSize(BTN_W, TAB_HEIGHT);
    m_addButton->setToolTip("Insert Sheet");
    m_addButton->setStyleSheet(
        "QToolButton { border: none; background: transparent; font-size: 16px; }"
        "QToolButton:hover { background: #C0C0C0; }"
    );
    connect(m_addButton, &QToolButton::clicked,
            this, &SheetTabBar::addSheetRequested);
}

int SheetTabBar::addTab(const QString& name)
{
    m_tabs.append(TabInfo{ name, QRect() });
    relayout();
    update();
    return m_tabs.size() - 1;
}

void SheetTabBar::removeTab(int index)
{
    if (index < 0 || index >= m_tabs.size()) return;
    m_tabs.removeAt(index);
    if (m_currentIndex >= m_tabs.size())
        m_currentIndex = m_tabs.size() - 1;
    relayout();
    update();
}

void SheetTabBar::renameTab(int index, const QString& name)
{
    if (index < 0 || index >= m_tabs.size()) return;
    m_tabs[index].name = name;
    relayout();
    update();
}

void SheetTabBar::setCurrentTab(int index)
{
    if (index < 0 || index >= m_tabs.size()) return;
    m_currentIndex = index;
    // Ensure visible
    if (m_scrollOffset > index)
        m_scrollOffset = index;
    relayout();
    update();
}

QString SheetTabBar::tabName(int index) const
{
    if (index < 0 || index >= m_tabs.size()) return {};
    return m_tabs[index].name;
}

QSize SheetTabBar::sizeHint() const
{
    return { 400, TAB_HEIGHT + 2 };
}

QSize SheetTabBar::minimumSizeHint() const
{
    return { 100, TAB_HEIGHT + 2 };
}

void SheetTabBar::relayout()
{
    QFontMetrics fm(font());
    int x = 2;

    for (int i = m_scrollOffset; i < m_tabs.size(); ++i) {
        int w = std::clamp(fm.horizontalAdvance(m_tabs[i].name) + 2 * TAB_PADDING,
                           MIN_TAB_W, MAX_TAB_W);
        m_tabs[i].rect = QRect(x, 1, w, TAB_HEIGHT);
        x += w;
    }

    // Place add button after last tab
    int addBtnX = x + 4;
    int rightEdge = width() - BTN_W * 2 - 4;
    m_addButton->move(std::min(addBtnX, rightEdge), 1);
    m_scrollLeft->move(width() - BTN_W * 2, 1);
    m_scrollRight->move(width() - BTN_W, 1);
}

void SheetTabBar::resizeEvent(QResizeEvent* e)
{
    QWidget::resizeEvent(e);
    relayout();
}

int SheetTabBar::tabAt(const QPoint& pt) const
{
    for (int i = 0; i < m_tabs.size(); ++i) {
        if (m_tabs[i].rect.contains(pt))
            return i;
    }
    return -1;
}

void SheetTabBar::paintEvent(QPaintEvent*)
{
    QPainter p(this);
    p.setRenderHint(QPainter::Antialiasing);

    QFont f = font();

    // Draw baseline
    p.setPen(QPen(QColor(0x217346), 2));
    p.drawLine(0, height() - 2, width(), height() - 2);

    for (int i = m_scrollOffset; i < m_tabs.size(); ++i) {
        const auto& tab = m_tabs[i];
        bool active = (i == m_currentIndex);
        QRect r = tab.rect;

        if (active) {
            // Active tab: white background, green top border
            p.setPen(Qt::NoPen);
            p.setBrush(Qt::white);
            p.drawRect(r);
            // Green top accent
            p.fillRect(r.left(), r.top(), r.width(), 3, QColor(0x217346));
        } else {
            // Inactive: light gray
            p.setPen(Qt::NoPen);
            p.setBrush(QColor(0xE8E8E8));
            p.drawRect(r);
            // Subtle border
            p.setPen(QColor(0xC0C0C0));
            p.drawRect(r.adjusted(0, 0, -1, -1));
        }

        // Tab name
        f.setBold(active);
        p.setFont(f);
        p.setPen(active ? QColor(0x217346) : QColor(0x444444));
        p.drawText(r.adjusted(TAB_PADDING/2, 0, -TAB_PADDING/2, 0),
                   Qt::AlignVCenter | Qt::AlignHCenter,
                   tab.name);
    }
}

void SheetTabBar::mousePressEvent(QMouseEvent* e)
{
    if (e->button() == Qt::LeftButton) {
        int idx = tabAt(e->pos());
        if (idx >= 0 && idx != m_currentIndex) {
            m_currentIndex = idx;
            update();
            emit tabActivated(idx);
        }
    }
    QWidget::mousePressEvent(e);
}

void SheetTabBar::mouseDoubleClickEvent(QMouseEvent* e)
{
    int idx = tabAt(e->pos());
    if (idx >= 0)
        emit renameSheetRequested(idx);
    QWidget::mouseDoubleClickEvent(e);
}

void SheetTabBar::contextMenuEvent(QContextMenuEvent* e)
{
    int idx = tabAt(e->pos());
    if (idx < 0) return;

    QMenu menu(this);
    auto* actRename = menu.addAction("Rename");
    auto* actInsert = menu.addAction("Insert Sheet...");
    menu.addSeparator();
    auto* actDelete = menu.addAction("Delete");

    auto* chosen = menu.exec(e->globalPos());
    if (chosen == actRename)
        emit renameSheetRequested(idx);
    else if (chosen == actInsert)
        emit addSheetRequested();
    else if (chosen == actDelete)
        emit deleteSheetRequested(idx);
}
