#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <ActiveQt/QAxObject>
#include "ExcelBase.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    QString getFileName();
    ~MainWindow();

private slots:
    void on_batchButton_clicked();

    void on_selectButton_clicked();

    void on_singleButton_clicked();

    void on_openButton_clicked();

    void on_clearLog_clicked();

    void on_dataCol_textEdited(const QString &arg1);

    void on_itemCol_textEdited(const QString &arg1);

private:
    Ui::MainWindow *ui;
    QScopedPointer<ExcelBase> m_xls;
    QList< QList<QVariant> > m_datas;
};



class QPosition{
public:
    int row = 1;
    int col = 1;
    QString cell = trans(col)+QString::number(row);
    QString trans(int num){
        QString ret="";
        QChar ch;
        while(num!=0) {
            int tmp = num %26;
            num/=26;
            //此处略微关键，当为0时，其实是26，也就是Z，
            //而且当你将0调整为26后，需要从数字中去除26代表的这个数
            if(tmp ==0) {
                tmp = 26;
                num-=1;
            }
            ch='A'+tmp-1;
            ret=ch+ret;
        }
        return ret;
    }
    QPosition(int a=1,int b=1):
        row(a),
        col(b),
        cell(trans(col)+QString::number(row)){}
    void set(int a,int b){
        row = a;
        col = b;
        cell = trans(col)+QString::number(row);
    }
};

#endif // MAINWINDOW_H
