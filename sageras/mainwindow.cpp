#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include "pshare.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->progressBar->setRange(0,100);
    ui->progressBar->setValue(0);

    thread1=new CopyThread(ui);
    connect(thread1,&CopyThread::message
            ,this,&MainWindow::receiveMessage);
    connect(thread1,&CopyThread::progress
            ,this,&MainWindow::progress);

    thread2=new ExcelBatchHandel(ui);
    connect(thread2,&ExcelBatchHandel::message
            ,this,&MainWindow::receiveMessage);
    connect(thread2,&ExcelBatchHandel::progress
            ,this,&MainWindow::progress);
}

MainWindow::~MainWindow()
{
    delete ui;
    delete thread2;
    delete thread1;
}

void MainWindow::on_openButton_clicked()
{
    QString path = ui->inputPath->text();
    QStringList strs = getPathFileNames(path);
    QListWidgetItem *aItem; //每一行是一个QListWidgetItem

    ui->listWidget->clear(); //清除项
    for (int i=0; i<strs.size(); i++)
    {
        aItem=new QListWidgetItem(); //新建一个项
        aItem->setText(strs[i]); //设置文字标签
        aItem->setCheckState(Qt::Unchecked); //设置为选中状态
        if (false) //可编辑, 设置flags
            aItem->setFlags(Qt::ItemIsSelectable | Qt::ItemIsEditable |Qt::ItemIsUserCheckable |Qt::ItemIsEnabled);
        else//不可编辑, 设置flags
            aItem->setFlags(Qt::ItemIsSelectable |Qt::ItemIsUserCheckable |Qt::ItemIsEnabled);
        ui->listWidget->addItem(aItem); //增加一个项
    }
    ui->hint->append(getSystemTime()+'\n'+"打开文件夹: "+path);
}

void MainWindow::on_selectButton_clicked()
{
    QString path=getOpenFileName();
    path = QDir::toNativeSeparators(path);
    ui->outputFile->setText(path);
    ui->hint->append(getSystemTime()+'\n'+"outputFile:"+path);
}

void MainWindow::on_batchButton_clicked()
{
    ui->progressBar->setValue(0);
    ui->copySheet->setEnabled(false);
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
    }else if(str=="batch_finish"){
        ui->batchButton->setEnabled(true);
    }
    else
        ui->hint->append(str);
}

void MainWindow::progress(int val)
{
    ui->progressBar->setValue(val);
}

void MainWindow::on_toolButton_clicked()
{
    int cnt=ui->listWidget->count();//项个数
    for (int i=0; i<cnt; i++)
    {
        QListWidgetItem *aItem=ui->listWidget->item(i);//获取一个项
        aItem->setCheckState(Qt::Checked);//设置为选中
    }
}

void MainWindow::on_toolButton_3_clicked()
{
    int cnt=ui->listWidget->count();//项个数
    for (int i=0; i<cnt; i++)
    {
        QListWidgetItem *aItem=ui->listWidget->item(i);//获取一个项
        aItem->setCheckState(Qt::Unchecked);//设置为选中
    }
}

void MainWindow::on_batchStop_clicked()
{
    thread2->stop();
    thread1->stop();
}

void MainWindow::on_toolButton_2_clicked()
{
    int cnt=ui->listWidget->count();//项个数
    for (int i=0; i<cnt; i++)
    {
        QListWidgetItem *aItem=ui->listWidget->item(i);//获取一个项
        Qt::CheckState flag = aItem->checkState();
        flag = flag==Qt::Checked?Qt::Unchecked:Qt::Checked;
        aItem->setCheckState(flag);//设置为选中
    }
}

void MainWindow::on_copySheet_clicked()
{
    ui->progressBar->setValue(0);
    ui->copySheet->setEnabled(false);
    ui->batchButton->setEnabled(false);
    thread1->start();
}
