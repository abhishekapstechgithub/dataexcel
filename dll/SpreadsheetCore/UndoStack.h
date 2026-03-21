#pragma once
#include <QStack>
#include <functional>

struct Command {
    std::function<void()> execute;
    std::function<void()> unexecute;
    QString description;
};

class UndoStack {
public:
    void push(Command cmd);
    void undo();
    void redo();
    bool canUndo() const;
    bool canRedo() const;
    void clear();
private:
    QStack<Command> m_undo;
    QStack<Command> m_redo;
    static constexpr int MAX_DEPTH = 200;
};
