#ifndef EXCELHANDLE_H
#define EXCELHANDLE_H

#include <QThread>
#include <QTime>
#include "mainwindow.h"

class ExcelHandel : public QThread
{
    Q_OBJECT
signals:
    void message(const QString& info);
    void progress(int present);
public:
    ExcelHandel(Ui::MainWindow *a,MainWindow* b);
    ~ExcelHandel();
    void closeThread();
    void run();
    void send(QString msg);
    virtual bool getdata();
    virtual bool deal();

private:
    volatile bool isStop;       //isStop是易失性变量，需要用volatile进行申明
    Ui::MainWindow *ui;
    MainWindow* p;

    QString inputFile="";
    QString outputFile="";
    QString sheet="";
    QString dataStart = "";
    QString itemStart = "";
    QString startRow = "";

    QString msg="";
    QTime t;

};

#endif // EXCELHANDLE_H
