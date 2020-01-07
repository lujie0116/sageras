#ifndef PSHARE_H
#define PSHARE_H

#include <ActiveQt/QAxObject>
#include <QDateTime>
#include <QMessageBox>
#include <QFileDialog>


//excel列进制转换
int stringToIntBy26Base(QString colName);
//获取系统时间
QString getSystemTime();
QString appendlog(QString msg);
QString getOpenFileName();
QStringList getPathFileNames(const QString &path);
bool fileExist(QString &file);
void getFileNametoTitle(QString &src,QString &num,QString &name);

//判断excel字段
bool isSubsection(QString str);
bool isOther(QString str);

void connectComponent(QAxObject* excel);
void excelFree(QAxObject* excel);

//处理excel
bool preProcess(QString path,QHash<QString,QHash<QString,QString>> &map,QAxObject* excel,QString &sheetName);
bool processFile(QString path,QAxObject* excel,QHash<QString,QHash<QString,QString>> &map,QString &dataStart,QString &itemStart,QString &startRow,QString &sheetName,QString inputFile);
bool readModule(QString path,QVariantList &headers);

//
bool readIni(QString path,QVariantList &headers);
void divideStrtoList(QString src,QChar tag,QVariantList &headers);

bool getListMap(const QVariantList &src, const QVariantList &target, QVariantList &map);
#endif // PSHARE_H
