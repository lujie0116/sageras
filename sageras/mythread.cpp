#include "mythread.h"
#include <QDebug>
#include <QMutex>
#include "pshare.h"

MyThread1::MyThread1(Ui::MainWindow *a,MainWindow* b):ui(a),p(b)
{
    isStop = false;
}

void MyThread1::closeThread()
{
    isStop = true;
}

void MyThread1::run()
{
    singleDeal(ui,p);
}

MyThread2::MyThread2(Ui::MainWindow *a,MainWindow* b):ui(a),p(b)
{
    isStop = false;
}

void MyThread2::closeThread()
{
    isStop = true;
}

void MyThread2::run()
{
    batchDeal(ui,p);
}
