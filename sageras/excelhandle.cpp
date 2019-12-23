#include "excelhandle.h"
#include <QDebug>
#include <QMutex>
#include "pshare.h"

ExcelHandel::ExcelHandel(Ui::MainWindow *a):ui(a)
{
    isStop = false;
}
ExcelHandel::~ExcelHandel(){
}

void ExcelHandel::closeThread()
{
    isStop = true;
}

void ExcelHandel::run()
{
    t.start();//将此时间设置为当前时间
//    QString str = QString("%1->%2,thread id:%3").arg(__FILE__).arg(__FUNCTION__).arg((unsigned int)QThread::currentThreadId());
//    send(str);
    if(!getdata()){
        emit message("single_finish");
        emit progress(100);
        return ;
    }
    if(!deal()){
        emit message("single_finish");
        emit progress(100);
        return ;
    }
    //    elapsed(): 返回自上次调用start()或restart()以来经过的毫秒数
    emit progress(100);
    msg="处理完成,花费时间为"+QString::number(t.elapsed())+"ms";
    send(msg);
    emit message("single_finish");
}

bool ExcelHandel::getdata(){
    if(ui->fileList->currentItem()==NULL){
        msg = "选择输入文件";
        send(msg);
        return false;
    }
    inputFile = ui->inputPath->text()+"\\"+ui->fileList->currentItem()->text();
    //检查文件是否存在
    if(!fileExist(inputFile)){
        msg = "输入文件不存在,请选择或检查路径";
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
    if(inputFile==outputFile){
        msg = "错误:输入输出为同一文件";
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

bool ExcelHandel::deal(){
    QHash<QString,QHash<QString,QString>> map;
    //连接控件
    QAxObject* excel = new QAxObject();
    connectComponent(excel);
    if(!preProcess(inputFile,map,excel,sheet)){
        msg="文件:"+inputFile+"读取出错";
        send(msg);
        excelFree(excel);
        return false;
    }

    bool result = processFile(outputFile,excel,map,dataStart,itemStart,startRow,sheet);
    excelFree(excel);
    if(result){
        msg = "文件:"+inputFile+"导入成功";
        send(msg);
        return true;
    }
    else{
        msg = "文件:"+inputFile+"导入失败";
        send(msg);
        return false;
    }
    return true;
}

void ExcelHandel::send(QString msg){
    emit message(appendlog(msg));
}
