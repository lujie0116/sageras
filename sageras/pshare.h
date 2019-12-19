#ifndef PSHARE_H
#define PSHARE_H

#include <ActiveQt/QAxObject>
#include <QDateTime>
#include <QMessageBox>
#include <QFileDialog>
#include "ui_mainwindow.h"

bool isSubsection(QString str);
bool isOther(QString str);
void connectComponent(QAxObject* excel);
QString getSystemTime();
QStringList getFileNames(const QString &path);
int stringToIntBy26Base(QString colName);
void preProcess(QString path,QMap<QString,QMap<QString,QString>> &map,QAxObject* excel);
bool processFile(QString path,QAxObject* excel,QMap<QString,QMap<QString,QString>> &map,QString &dataStart,QString &itemStart,QString &startRow);
void singleDeal(Ui::MainWindow *ui,MainWindow* p);
void batchDeal(Ui::MainWindow *ui,MainWindow* p);

#endif // PSHARE_H
