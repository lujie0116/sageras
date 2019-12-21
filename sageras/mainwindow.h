#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <ActiveQt/QAxObject>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_batchButton_clicked();

    void on_selectButton_clicked();

    void on_singleButton_clicked();

    void on_openButton_clicked();

    void on_clearLog_clicked();

    void on_dataCol_textEdited(const QString &arg1);

    void on_itemCol_textEdited(const QString &arg1);

    void progress(int val);

    void receiveMessage(const QString& str);

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
