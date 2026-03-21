#include "UndoStack.h"
#include <QStack>

void UndoStack::push(Command cmd) {
    cmd.execute();
    m_undo.push(cmd);
    m_redo.clear();
    if (m_undo.size() > MAX_DEPTH) {
        // Remove oldest: rebuild without bottom
        QStack<Command> tmp;
        while (!m_undo.isEmpty()) tmp.push(m_undo.pop());
        tmp.pop(); // discard oldest
        while (!tmp.isEmpty()) m_undo.push(tmp.pop());
    }
}

void UndoStack::undo() {
    if (m_undo.isEmpty()) return;
    Command cmd = m_undo.pop();
    cmd.unexecute();
    m_redo.push(cmd);
}

void UndoStack::redo() {
    if (m_redo.isEmpty()) return;
    Command cmd = m_redo.pop();
    cmd.execute();
    m_undo.push(cmd);
}

bool UndoStack::canUndo() const { return !m_undo.isEmpty(); }
bool UndoStack::canRedo() const { return !m_redo.isEmpty(); }
void UndoStack::clear() { m_undo.clear(); m_redo.clear(); }
