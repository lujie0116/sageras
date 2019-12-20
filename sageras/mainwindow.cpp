#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QElapsedTimer>
#include <qDebug>
#include "mythread.h"
#include "pshare.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

QString MainWindow::getFileName(){
    return QFileDialog::getOpenFileName(this,tr("文件对话框"),"E:\\download",tr("excel文件(*.xlsx);;""文件(*)"));
}

void MainWindow::on_openButton_clicked()
{
    ui->fileList->clear();
    QString path = ui->inputPath->text();
    QStringList strs = getFileNames(path);
    ui->fileList->addItems(strs);
    ui->hint->append(getSystemTime()+'\n'+"打开文件夹: "+path);
}

void MainWindow::on_selectButton_clicked()
{
    QString path=getFileName();
    path = QDir::toNativeSeparators(path);
    ui->outputFile->setText(path);
    ui->hint->append(getSystemTime()+'\n'+"outputFile:"+path);
}

void MainWindow::on_singleButton_clicked()
{
//    MyThread1 *thread=new MyThread1(ui,this);
//    thread->start();
    //QString xlsFile = QFileDialog::getOpenFileName(this,QString(),QString(),"excel(*.xls *.xlsx)");
    QString xlsFile = ui->inputPath->text()+"\\"+ui->fileList->currentItem()->text();
//    if(xlsFile.isEmpty())
//        return;
    QElapsedTimer timer;
    timer.start();
    if(m_xls.isNull())
        m_xls.reset(new ExcelBase);
    m_xls->open(xlsFile);
    qDebug()<<"open cost:"<<timer.elapsed()<<"ms";timer.restart();
    m_xls->setCurrentSheet(1);
    m_xls->readAll(m_datas);
    qDebug()<<"read data cost:"<<timer.elapsed()<<"ms";timer.restart();
//    QVariantListListModel* md = qobject_cast<QVariantListListModel*>(ui->tableView->model());
//    if(md)
//    {
//        md->updateData();
//    }
//    qDebug()<<"show data cost:"<<timer.elapsed()<<"ms";timer.restart();
}

void MainWindow::on_batchButton_clicked()
{
    MyThread2 *thread=new MyThread2(ui,this);
    thread->start();
}

void MainWindow::on_clearLog_clicked()
{
    QString blank="";
    ui->hint->setText(blank);
}

void MainWindow::on_dataCol_textEdited(const QString &arg1)
{
    QString str=arg1.toUpper();
    ui->dataCol->setText(str);
}

void MainWindow::on_itemCol_textEdited(const QString &arg1)
{
    QString str=arg1.toUpper();
    ui->itemCol->setText(str);
}
