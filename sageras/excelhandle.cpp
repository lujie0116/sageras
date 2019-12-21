#include "excelhandle.h"
#include <QDebug>
#include <QMutex>
#include "pshare.h"

ExcelHandel::ExcelHandel(Ui::MainWindow *a,MainWindow* b):ui(a),p(b)
{
    isStop = false;
}

void ExcelHandel::closeThread()
{
    isStop = true;
}

void ExcelHandel::run()
{
    singleDeal(ui,p);
}
