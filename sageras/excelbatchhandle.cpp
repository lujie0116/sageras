#include "excelbatchhandle.h"
#include <QDebug>
#include <QMutex>
#include "pshare.h"
#include "qposition.h"

ExcelBatchHandel::ExcelBatchHandel(Ui::MainWindow *a):ui(a)
{
    isStop = false;
    inputFiles.clear();
}

void ExcelBatchHandel::closeThread()
{
    isStop = true;
}

void ExcelBatchHandel::run()
{
    isStop = false;
    t.start();//将此时间设置为当前时间
//    QString str = QString("%1->%2,thread id:%3").arg(__FILE__).arg(__FUNCTION__).arg((unsigned int)QThread::currentThreadId());
//    send(str);
    if(!getdata()){
        emit message("batch_finish");
        return ;
    }
    if(!deal()){
        emit message("batch_finish");
        return ;
    }
    //    elapsed(): 返回自上次调用start()或restart()以来经过的毫秒数
    isStop = false;
    msg="处理结束,花费时间为"+QString::number(t.elapsed())+"ms";
    send(msg);
    emit message("batch_finish");
}

void ExcelBatchHandel::stop()
{
    isStop = true;
}

bool ExcelBatchHandel::getdata(){
    inputPath = ui->inputPath->text();
    inputFiles.clear();

    int cnt=ui->listWidget->count();//项个数
    for (int i=0; i<cnt; i++)
    {
        QListWidgetItem *aItem=ui->listWidget->item(i);//获取一个项
        if(aItem->checkState()==Qt::Checked)
            inputFiles.append(aItem->text());
    }

    if(inputFiles.size()==0){
        msg = "请选择输入文件";
        send(msg);
        return false;
    }

    outputFile = ui->outputFile->text();
    if(outputFile==""){
        msg = "选择输出文件";
        send(msg);
        return false;
    }
    if(!fileExist(outputFile)){
        msg = "输出文件不存在,请检查";
        send(msg);
        return false;
    }

    sheet = ui->sheetName->text();
    if(sheet==""){
        msg = "工作表不能为空.";
        send(msg);
        return false;
    }
    dataStart = ui->dataCol->text();
    itemStart= ui->itemCol->text();
    startRow = ui->startRow->text();
    return true;
}

bool ExcelBatchHandel::deal(){
    //连接控件
    QAxObject* excel = new QAxObject();
    connectComponent(excel);
    int successcnt=0;
    int failcnt=0;
    int finishcnt=0;
    int col=stringToIntBy26Base(dataStart);
    int row=startRow.toInt();

    QPosition idx(row,col);

    for(int i=0;i<inputFiles.size();i++){
        QHash<QString,QHash<QString,QString>> map;
        QString inputFile=inputPath+'\\'+inputFiles[i];

        if(inputFile==outputFile){
            failcnt++;
            msg = "文件:"+inputFile+"为输出文件\n"+"成功"+QString::number(successcnt)+"个,失败"+QString::number(failcnt)+"个,剩余"+QString::number(inputFiles.size()-i-1)+"个";
            send(msg);
            finishcnt=successcnt+failcnt;
            emit progress(((float)finishcnt / inputFiles.size()) * 100);
            continue;
        }
        if(!preProcess(inputFile,map,excel,sheet)){
            failcnt++;
            msg = "文件:"+inputFile+"读取出错\n"+"成功"+QString::number(successcnt)+"个,失败"+QString::number(failcnt)+"个,剩余"+QString::number(inputFiles.size()-i-1)+"个";
            send(msg);
            finishcnt=successcnt+failcnt;
            emit progress(((float)finishcnt / inputFiles.size()) * 100);
            continue;
        }
        dataStart=idx.trans(col);
        if(processFile(outputFile,excel,map,dataStart,itemStart,startRow,sheet)){
            successcnt++;
            msg = "文件:"+inputFile+"导入成功\n"+"成功"+QString::number(successcnt)+"个,失败"+QString::number(failcnt)+"个,剩余"+QString::number(inputFiles.size()-i-1)+"个";
            send(msg);
        }
        else{
            failcnt++;
            msg = "文件:"+inputFile+"导入出错\n"+"成功"+QString::number(successcnt)+"个,失败"+QString::number(failcnt)+"个,剩余"+QString::number(inputFiles.size()-i-1)+"个";
            send(msg);
        }
        col++;
        finishcnt=successcnt+failcnt;
        emit progress(((float)finishcnt / inputFiles.size()) * 100);
        if(isStop) {
            send("停止");
            break;
        }
    }
    excelFree(excel);
    return true;
}

void ExcelBatchHandel::send(QString msg){
    emit message(appendlog(msg));
}
