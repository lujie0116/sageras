#ifndef PSHARE_H
#define PSHARE_H

#include <ActiveQt/QAxObject>
#include <QDateTime>
#include <QMessageBox>
#include <QFileDialog>
#include "ui_mainwindow.h"
#include "mainwindow.h"


//excel列进制转换
int stringToIntBy26Base(QString colName);
//获取系统时间
QString getSystemTime();
QString getOpenFileName();
QStringList getPathFileNames(const QString &path);

//判断excel字段
bool isSubsection(QString str);
bool isOther(QString str);

void connectComponent(QAxObject* excel);

//处理excel
void preProcess(QString path,QMap<QString,QMap<QString,QString>> &map,QAxObject* excel);
bool processFile(QString path,QAxObject* excel,QMap<QString,QMap<QString,QString>> &map,QString &dataStart,QString &itemStart,QString &startRow);

void singleDeal(Ui::MainWindow *ui,MainWindow* p);
void batchDeal(Ui::MainWindow *ui,MainWindow* p);

#endif // PSHARE_H
