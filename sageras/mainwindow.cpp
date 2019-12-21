#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include "excelhandle.h"
#include "excelbatchhandle.h"
#include "pshare.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->progressBar->setRange(0,100);
    ui->progressBar->setValue(0);
}

MainWindow::~MainWindow()
{
    delete ui;
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
    ExcelHandel *thread=new ExcelHandel(ui,this);
    thread->start();
}

void MainWindow::on_batchButton_clicked()
{
    ExcelBatchHandel *thread=new ExcelBatchHandel(ui,this);
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
