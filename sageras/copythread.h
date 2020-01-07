#ifndef COPYTHREAD_H
#define COPYTHREAD_H


#include <QThread>
#include <QTime>
#include "ui_mainwindow.h"

class CopyThread : public QThread
{
    Q_OBJECT
signals:
    void message(const QString& info);
    void progress(int present);
public:
    CopyThread(Ui::MainWindow *a);
    void closeThread();
    virtual void run();
    void stop();
    void send(QString msg);
    virtual bool getdata();
    virtual bool deal();

private:
    volatile bool isStop;       //isStop是易失性变量，需要用volatile进行申明
    Ui::MainWindow *ui;

    QStringList inputFiles;
    QString inputPath="";
    QString outputFile="";
    QString sheet="";
    QString dataStart = "";
    QString itemStart = "";
    QString startRow = "";

    QString msg="";
    QTime t;
};

#endif // COPYTHREAD_H
