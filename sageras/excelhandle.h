#ifndef EXCELHANDLE_H
#define EXCELHANDLE_H

#include <QThread>
#include <QTime>
#include "ui_mainwindow.h"

class ExcelHandel : public QThread
{
    Q_OBJECT
signals:
    void message(const QString& info);
    void progress(int present);
public:
    ExcelHandel(Ui::MainWindow *a);
    ~ExcelHandel();
    void closeThread();
    void run();
    void send(QString msg);
    virtual bool getdata();
    virtual bool deal();

private:
    volatile bool isStop;       //isStop是易失性变量，需要用volatile进行申明
    Ui::MainWindow *ui;

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
