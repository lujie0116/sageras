#include "excelbatchhandle.h"
#include <QDebug>
#include <QMutex>
#include "pshare.h"

ExcelBatchHandel::ExcelBatchHandel(Ui::MainWindow *a,MainWindow* b):ui(a),p(b)
{
    isStop = false;
}

void ExcelBatchHandel::closeThread()
{
    isStop = true;
}

void ExcelBatchHandel::run()
{
    batchDeal(ui,p);
}
