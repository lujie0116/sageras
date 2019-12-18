#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <ActiveQt/QAxObject>
#include <QDateTime>

bool isSubsection(QString str){
    QString filter="--";
    int idx=str.indexOf(filter);
    return idx!=-1;
}
bool isOther(QString str){
    QString filter="其他";
    int idx=str.indexOf(filter);
    return idx==0;
}

//连接控件
void connectComponent(QAxObject* excel){
    excel->setControl("Excel.Application");  // 连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", "false"); // 不显示窗体
    excel->setProperty("DisplayAlerts", false);  // 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示
}

QString getSystemTime()
{
    QDateTime current_date_time = QDateTime::currentDateTime();
    QString current_time = current_date_time.toString("hh:mm:ss.zzz ");

    return current_time;
}

QStringList getFileNames(const QString &path)
{
    QDir dir(path);
    QStringList nameFilters;
    nameFilters << "*.xlsx";
    QStringList files = dir.entryList(nameFilters, QDir::Files|QDir::Readable, QDir::Name);
    return files;
}
int stringToIntBy26Base(QString colName){
    int column = 0;
    int strLen = colName.length();
    for (int i=0,j=1; i<strLen; i++,j*=26) {
        int temp = (int)(colName.toUpper().at(strLen - 1 - i).toLatin1() - 64);
        column = column + temp * j;
    }
    return column;
}
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

void MainWindow::preProcess(QString path,QMap<QString,QMap<QString,QString>> &map,QAxObject* excel){
    if(path=="") return ;

    // step2: 打开工作簿
    QAxObject* workbooks = excel->querySubObject("WorkBooks"); // 获取工作簿集合

    //————————————————按文件路径打开文件————————————————————
    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", path);

    // 获取打开的excel文件中所有的工作sheet
    QAxObject* worksheets = workbook->querySubObject("Sheets");


    //—————————————————Excel文件中表的个数:——————————————————
    int iWorkSheet = worksheets->property("Count").toInt();
    QString sheetName;
    int curSheet;
    for(curSheet=1;curSheet<=iWorkSheet;curSheet++){
        sheetName=worksheets->querySubObject("Item(int)", curSheet)->property("Name").toString();
        if(sheetName=="G1-1") break;
    }

    if(sheetName!="G1-1") return;

    // ————————————————获取第n个工作表 querySubObject("Item(int)", n);——————————
    QAxObject * worksheet = worksheets->querySubObject("Item(int)", curSheet);//本例获取第一个，最后参数填1

    QAxObject* usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
    int intRowStart = usedrange->property("Row").toInt(); // 起始行数
    int intColStart = usedrange->property("Column").toInt();  // 起始列数

    QAxObject *rows, *columns;
    rows = usedrange->querySubObject("Rows");  // 行
    columns = usedrange->querySubObject("Columns");  // 列

    int intRow = rows->property("Count").toInt(); // 行数
    int intCol = columns->property("Count").toInt();  // 列数

    QString X="";
    QAxObject* cellX = NULL; //获取单元格
    QString str="";
    QPosition itemPosition(1,30);
    QPosition moneyPosition;
    bool item_flag=false,total_flag=false;;

    //查找项目和金额的位置
    for(int i=intRowStart;i<=intRow;i++){
        for(int j=intColStart;j<=intCol;j++){
            QPosition X(i,j);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)", X.cell);
            str = cellX->dynamicCall("Value2()").toString();
            if(str=="项目"){
                itemPosition.set(X.row,X.col);
                item_flag=true;
            }
            if(item_flag&&str=="金额"){
                moneyPosition.set(X.row,X.col);
                total_flag=true;
            }
        }
    }
    if(!total_flag) return ;

    QMap<QString,QString> subMap;
    QString section, count;
    QString other="=0";
    for(int i=itemPosition.row+1;i<=intRow;i++){
        QPosition item(i,itemPosition.col),money(i,moneyPosition.col);
        cellX = worksheet->querySubObject("Range(QVariant, QVariant)", item.cell);
        str = cellX->dynamicCall("Value2()").toString();
        cellX = worksheet->querySubObject("Range(QVariant, QVariant)", money.cell);
        count=cellX->dynamicCall("Value2()").toString();

        //两层hashmap保存
        section=str;
        if(section=="") continue;
        subMap.clear();

        if(isOther(section)){
            //todo::其他情况
            if(count!="")
                other+="+"+count;
            subMap.insert("--",other);
            map.insert("其他",subMap);
            continue;
        }

        //判断有没有子项目
        QPosition next_item(i+1,itemPosition.col),next_money(i+1,moneyPosition.col);
        cellX = worksheet->querySubObject("Range(QVariant, QVariant)", next_item.cell);
        str = cellX->dynamicCall("Value2()").toString();

        if(!isSubsection(str)){
            subMap.insert("--",count);
            map.insert(section,subMap);
            continue;
        }

        cellX = worksheet->querySubObject("Range(QVariant, QVariant)", next_money.cell);
        count=cellX->dynamicCall("Value2()").toString();

        while(isSubsection(str)){
            subMap.insert(str,count);
            i++;
            next_item.set(next_item.row+1,next_item.col);
            next_money.set(next_money.row+1,next_money.col);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)", next_item.cell);
            str = cellX->dynamicCall("Value2()").toString();
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)", next_money.cell);
            count=cellX->dynamicCall("Value2()").toString();
        }
        map.insert(section,subMap);
    }
//    workbook->dynamicCall("Save()", true);
    //关闭文件
//    workbook->dynamicCall("Close(Boolean)", true);
//    excel->dynamicCall("Quit()");
//    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");
}

bool processFile(QString path,QAxObject* excel,QMap<QString,QMap<QString,QString>> &map,QString &dataStart,QString &itemStart,QString &startRow){
    bool ret=false;
    if(path=="") return ret;

    // step2: 打开工作簿
    QAxObject* workbooks = excel->querySubObject("WorkBooks"); // 获取工作簿集合

    //————————————————按文件路径打开文件————————————————————
    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", path);

    // 获取打开的excel文件中所有的工作sheet
    QAxObject * worksheets = workbook->querySubObject("WorkSheets");


    //—————————————————Excel文件中表的个数:——————————————————

    QAxObject * worksheet = worksheets->querySubObject("Item(int)", 1);//本例获取第一个，最后参数填1

    QAxObject* usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
    int RowStart = startRow.toInt(); // 起始行数
    int dataCol = stringToIntBy26Base(dataStart);
    int itemCol = stringToIntBy26Base(itemStart);

    QAxObject *rows, *columns;
    rows = usedrange->querySubObject("Rows");  // 行
    columns = usedrange->querySubObject("Columns");  // 列

    int intRow = rows->property("Count").toInt(); // 行数

    QAxObject* cellX = NULL; //获取单元格
    QString str="";
    QPosition moneyPosition(RowStart,dataCol);

    QString section="";
    QPosition targetPosition;
    for(int i=RowStart;i<=intRow;i++){
        QPosition X(i,itemCol);
        cellX = worksheet->querySubObject("Range(QVariant, QVariant)", X.cell);
        str = cellX->dynamicCall("Value2()").toString();

        if(!isSubsection(str)){
            section=str;

            QPosition next_item(i+1,1);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)", next_item.cell);
            str = cellX->dynamicCall("Value2()").toString();

            //if(isOther(str)) continue;

            if(isSubsection(str)) continue;

            targetPosition.set(i,moneyPosition.col);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)", targetPosition.cell);
            cellX->dynamicCall("SetValue(const QVariant&)",QVariant(map[section]["--"]));
        }else{
            targetPosition.set(i,moneyPosition.col);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)", targetPosition.cell);
            cellX->dynamicCall("SetValue(const QVariant&)",QVariant(map[section][str]));
        }
    }
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");
    return true;

}

void MainWindow::on_openButton_clicked()
{
    ui->fileList->clear();
    QString path = ui->inputPath->text();
    QStringList strs = getFileNames(path);
    ui->fileList->addItems(strs);
    ui->hint->append(getSystemTime()+'\n'+"打开文件夹: "+path);

    QPosition as(1,ui->startRow->text().toInt());
    ui->hint->append(as.cell+'\n');
}

void MainWindow::on_selectButton_clicked()
{
    QString path=getFileName();
    ui->outputFile->setText(path);
    ui->hint->append(getSystemTime()+'\n'+"outputFile:"+path);
}

void MainWindow::on_singleButton_clicked()
{
    QString inputFile = ui->inputPath->text()+"\\"+ui->fileList->currentItem()->text();
    QString outputFile = ui->outputFile->text();

    QMap<QString,QMap<QString,QString>> map;
    //连接控件
    QAxObject* excel = new QAxObject(this);
    connectComponent(excel);
    preProcess(inputFile,map,excel);

    QString dataStart = ui->dataCol->text();
    QString itemStart= ui->itemCol->text();
    QString startRow = ui->startRow->text();

    bool result = processFile(outputFile,excel,map,dataStart,itemStart,startRow);

    excel->dynamicCall("Quit()");
    if (excel)
    {
        delete excel;
        excel = NULL;
    }
    if(result)
        ui->hint->append(getSystemTime()+'\n'+"inputFile: "+inputFile+'\n'+"outputFile: "+outputFile);
    else
        ui->hint->append(getSystemTime()+'\n'+"inputFile: "+inputFile+"导入出错");
}

void MainWindow::on_batchButton_clicked()
{
    QString path = ui->inputPath->text();
    QStringList strs = getFileNames(path);
    QString outputFile = ui->outputFile->text();
    ui->hint->append(getSystemTime()+'\n'+"批处理: "+path);
    QString dataStart = ui->dataCol->text();
    QString itemStart= ui->itemCol->text();
    QString startRow = ui->startRow->text();

    int col=stringToIntBy26Base(dataStart);
    int row=startRow.toInt();

    QPosition idx(row,col);
    QAxObject* excel = new QAxObject(this);
    connectComponent(excel);
    for(int i=0;i<strs.size();i++){
        QString inputFile=path+'\\'+strs[i];
        QMap<QString,QMap<QString,QString>> map;
        //连接控件
        preProcess(inputFile,map,excel);

        dataStart=idx.trans(col);
        bool result = processFile(outputFile,excel,map,dataStart,itemStart,startRow);
        col++;
        if(result)
            ui->hint->append(getSystemTime()+'\n'+"inputFile: "+inputFile+'\n'+"outputFile: "+outputFile);
        else
            ui->hint->append(getSystemTime()+'\n'+"inputFile: "+inputFile+"导入出错");
    }

    excel->dynamicCall("Quit()");
    if (excel)
    {
        delete excel;
        excel = NULL;
    }

    ui->hint->append(getSystemTime()+'\n'+"批处理结束");
}


