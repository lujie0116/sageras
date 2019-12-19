#ifndef MYTHREAD_H
#define MYTHREAD_H

#include <QThread>
#include "mainwindow.h"

class MyThread1 : public QThread
{
public:
    MyThread1(Ui::MainWindow *a,MainWindow* b);
    void closeThread();

protected:
    virtual void run();

private:
    volatile bool isStop;       //isStop是易失性变量，需要用volatile进行申明
    Ui::MainWindow *ui;
    MainWindow* p;
};

class MyThread2 : public QThread
{
public:
    MyThread2(Ui::MainWindow *a,MainWindow* b);
    void closeThread();

protected:
    virtual void run();

private:
    volatile bool isStop;       //isStop是易失性变量，需要用volatile进行申明
    Ui::MainWindow *ui;
    MainWindow* p;
};

#endif // MYTHREAD_H
