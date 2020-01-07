#include <ActiveQt/QAxObject>
#include <QDateTime>
#include <QMessageBox>
#include <QFileDialog>
#include <QSettings>
#include "pshare.h"
#include "qposition.h"

//进制转换
int stringToIntBy26Base(QString colName){
    int column = 0;
    int strLen = colName.length();
    for (int i=0,j=1; i<strLen; i++,j*=26) {
        int temp = (int)(colName.toUpper().at(strLen - 1 - i).toLatin1() - 64);
        column = column + temp * j;
    }
    return column;
}
//获取系统时间
QString getSystemTime()
{
    QDateTime current_date_time = QDateTime::currentDateTime();
    QString current_time = current_date_time.toString("hh:mm:ss.zzz ");
    return current_time;
}

bool fileExist(QString &file){
    QFileInfo fileInfo(file);
    return fileInfo.exists();
}

void getFileNametoTitle(QString &str,QString &num,QString &name){
    num=str.mid(0,3);
    int start=str.lastIndexOf('-')+1;
    name=str.mid(start,str.size()-start);
}

QString appendlog(QString msg){
    return getSystemTime()+'\n'+msg;
}

//打开对话框选择文件
QString getOpenFileName(){
    return QFileDialog::getOpenFileName(NULL,"文件对话框","E:\\download","excel文件(*.xlsx);;""文件(*)");
}
//输出path下excel文件
QStringList getPathFileNames(const QString &path)
{
    QDir dir(path);
    QStringList nameFilters;
    nameFilters << "*.xlsx";
    QStringList files = dir.entryList(nameFilters, QDir::Files|QDir::Readable, QDir::Name);
    return files;
}

bool isSubsection(QString str){
    QString filter="--";
    int idx=str.indexOf(filter);
    return idx!=-1;
}

bool isOther(QString str){
    QStringList strs={"其他","其它"};
    for(int i=0;i<strs.size();i++){
        int idx=str.indexOf(strs[i]);
        if(idx!=-1) return true;
    }
    return false;
}

//连接控件
void connectComponent(QAxObject* excel){
    excel->setControl("Excel.Application");  // 连接Excel控件
    excel->dynamicCall("SetVisible (bool Visible)", "false"); // 不显示窗体
    excel->setProperty("DisplayAlerts", false);  // 不显示任何警告信息。如果为true, 那么关闭时会出现类似"文件已修改，是否保存"的提示
}

void excelFree(QAxObject* excel){
    excel->dynamicCall("Quit()");
    if (excel)
    {
        delete excel;
        excel = NULL;
    }
}

bool preProcess(QString path,QHash<QString,QHash<QString,QString>> &map,QAxObject* excel,QString &sheetName){
    if(path=="") return false;

    // step2: 打开工作簿
    QAxObject* workbooks = excel->querySubObject("WorkBooks"); // 获取工作簿集合
    if(workbooks==NULL) return false;

    //————————————————按文件路径打开文件————————————————————
    QAxObject* workbook = workbooks->querySubObject("Open(QString, QVariant)", path, 0);
    if(workbook==NULL) return false;

    // 获取打开的excel文件中所有的工作sheet
    QAxObject* worksheets = workbook->querySubObject("Sheets");
    if(worksheets==NULL) return false;

    //—————————————————Excel文件中表的个数:——————————————————
    int iWorkSheet = worksheets->property("Count").toInt();
    QString curSheetName;
    int curSheet;
    for(curSheet=1;curSheet<=iWorkSheet;curSheet++){
        curSheetName=worksheets->querySubObject("Item(int)", curSheet)->property("Name").toString();
        if(curSheetName==sheetName) break;
    }

    if(curSheetName!=sheetName) return false;

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
                break;
            }
        }
        if(total_flag) break;
    }
    if(!total_flag) return false;

    QHash<QString,QString> subMap;
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

        if(section=="合计")
            break;

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
    return true;
}

bool processFile(QString path,QAxObject* excel,QHash<QString,QHash<QString,QString>> &map,QString &dataStart,QString &itemStart,QString &startRow,QString &sheetName,QString inputFile){
    if(path=="") return false;

    // step2: 打开工作簿
    QAxObject* workbooks = excel->querySubObject("WorkBooks"); // 获取工作簿集合
    if(workbooks==NULL) return false;

    //————————————————按文件路径打开文件————————————————————
    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", path);
    if(workbook==NULL) return false;

    // 获取打开的excel文件中所有的工作sheet
    QAxObject * worksheets = workbook->querySubObject("WorkSheets");
    if(worksheets==NULL) return false;

    //—————————————————Excel文件中表的个数:——————————————————

    int iWorkSheet = worksheets->property("Count").toInt();
    QString curSheetName;
    int curSheet;
    for(curSheet=1;curSheet<=iWorkSheet;curSheet++){
        curSheetName=worksheets->querySubObject("Item(int)", curSheet)->property("Name").toString();
        if(curSheetName==sheetName) break;
    }

    if(curSheetName!=sheetName) return false;

    // ————————————————获取第n个工作表 querySubObject("Item(int)", n);——————————
    QAxObject * worksheet = worksheets->querySubObject("Item(int)", curSheet);//本例获取第一个，最后参数填1

    if(worksheet==NULL) return false;

    QAxObject* usedrange = worksheet->querySubObject("UsedRange"); // sheet范围
    int RowStart = startRow.toInt(); // 起始行数
    int dataCol = stringToIntBy26Base(dataStart);
    int itemCol = stringToIntBy26Base(itemStart);

    QAxObject *rows;
    rows = usedrange->querySubObject("Rows");  // 行
    //columns = usedrange->querySubObject("Columns");  // 列

    int intRow = rows->property("Count").toInt(); // 行数

    QAxObject* cellX = NULL; //获取单元格
    QString str="";
    QPosition moneyPosition(RowStart,dataCol);

    QString section="";
    QPosition targetPosition;

    QPosition file(1,dataCol);
    cellX = worksheet->querySubObject("Range(QVariant, QVariant)", file.cell);
    cellX->dynamicCall("SetValue(conts QVariant&)", inputFile);

    //获取S10和长沙美东
    int s=inputFile.indexOf(' ');
    QString shopNum="";
    if(shopNum!=-1){
        QPosition X(2,dataCol);
        cellX = worksheet->querySubObject("Range(QVariant, QVariant)", X.cell);
        shopNum=inputFile.mid(0,s);
        cellX->dynamicCall("SetValue(conts QVariant&)", shopNum);
    }


    s=inputFile.lastIndexOf('-');
    QString shopName="";
    if(shopName!=-1){
        QPosition X(3,dataCol);
        cellX = worksheet->querySubObject("Range(QVariant, QVariant)", X.cell);
        shopName=inputFile.mid(s+1,inputFile.lastIndexOf('.')-s-1);
        cellX->dynamicCall("SetValue(conts QVariant&)", shopName);
    }


    //存数值
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
            if(map[section].count("--")!=0)
                cellX->dynamicCall("SetValue(const QVariant&)",QVariant(map[section]["--"]));
        }else{
            targetPosition.set(i,moneyPosition.col);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)", targetPosition.cell);            
            if(map[section].count(str)!=0)
                cellX->dynamicCall("SetValue(const QVariant&)",QVariant(map[section][str]));
        }
    }
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");
    return true;

}

bool readModule(QString path,QVariantList &headers){

}

bool readIni(QString path,QVariantList &headers){
    QSettings setting("config.ini",QSettings::IniFormat);
    setting.beginGroup("config");//beginGroup与下面endGroup 相对应，“config”是标记
    setting.setIniCodec("UTF-8");
    //读取ini
    QString str=setting.value("headers").toString();
    setting.endGroup();
    QChar ch='\t';
    divideStrtoList(str,ch,headers);
    return true;

}

void divideStrtoList(QString src,QChar tag,QVariantList &headers){
    int idx=0;
    while(src.size()!=0){
        idx=src.indexOf(tag);
        headers.append(src.mid(0,idx));
        if(idx==-1)
            break;
        src=src.mid(idx+1,src.size()-idx-1);
    }
}

bool getListMap(const QVariantList &src,const QVariantList &target, QVariantList &map){
    for(int i=0;i<src.size();i++){
        for(int j=0;j<target.size();j++){
            if(src[i]==target[j]){
                map.append(j);
                break;
            }else
                return false;
        }
    }
   return true;
}

