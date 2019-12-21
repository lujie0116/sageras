#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include "excelhandle.h"
#include "excelbatchhandle.h"
#include "pshare.h"

ExcelHandel* thread1=NULL;
ExcelBatchHandel* thread2=NULL;

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->progressBar->setRange(0,100);
    ui->progressBar->setValue(0);

    thread1=new ExcelHandel(ui,this);
    connect(thread1,&ExcelHandel::message
            ,this,&MainWindow::receiveMessage);
    connect(thread1,&ExcelHandel::progress
            ,this,&MainWindow::progress);

    thread2=new ExcelBatchHandel(ui,this);
    connect(thread2,&ExcelBatchHandel::message
            ,this,&MainWindow::receiveMessage);
    connect(thread2,&ExcelBatchHandel::progress
            ,this,&MainWindow::progress);
}

MainWindow::~MainWindow()
{
    delete ui;
    delete thread1;
    delete thread2;
}

void MainWindow::on_openButton_clicked()
{
    ui->fileList->clear();
    QString path = ui->inputPath->text();
    QStringList strs = getPathFileNames(path);
    ui->fileList->addItems(strs);
    ui->hint->append(getSystemTime()+'\n'+"打开文件夹: "+path);
}

void MainWindow::on_selectButton_clicked()
{
    QString path=getOpenFileName();
    path = QDir::toNativeSeparators(path);
    ui->outputFile->setText(path);
    ui->hint->append(getSystemTime()+'\n'+"outputFile:"+path);
}

void MainWindow::on_singleButton_clicked()
{
    ui->progressBar->setValue(0);
    ui->singleButton->setEnabled(false);
    ui->batchButton->setEnabled(false);
    thread1->start();
}

void MainWindow::on_batchButton_clicked()
{
    ui->progressBar->setValue(0);
    ui->singleButton->setEnabled(false);
    ui->batchButton->setEnabled(false);
    thread2->start();
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

void MainWindow::receiveMessage(const QString &str)
{
    if(str=="single_finish"){
        ui->batchButton->setEnabled(true);
        ui->singleButton->setEnabled(true);
    }else if(str=="batch_finish"){
        ui->batchButton->setEnabled(true);
        ui->singleButton->setEnabled(true);
    }
    else
        ui->hint->append(str);
}

void MainWindow::progress(int val)
{
    ui->progressBar->setValue(val);
}
