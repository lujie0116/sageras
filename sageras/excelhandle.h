#ifndef EXCELHANDLE_H
#define EXCELHANDLE_H

#include <QThread>
#include "mainwindow.h"

class ExcelHandel : public QThread
{
public:
    ExcelHandel(Ui::MainWindow *a,MainWindow* b);
    void closeThread();

protected:
    virtual void run();

private:
    volatile bool isStop;       //isStop是易失性变量，需要用volatile进行申明
    Ui::MainWindow *ui;
    MainWindow* p;
};

#endif // EXCELHANDLE_H
