#include "qposition.h"

QPosition::QPosition(int a,int b)
    : row(a), col(b), cell(trans(col)+QString::number(row))
{
}

QString QPosition::trans(int num){
    QString ret="";
    QChar ch;
    while(num!=0) {
        int tmp = num %26;
        num/=26;
        //此处略微关键，当为0时，其实是26，也就是Z，
        //而且当你将0调整为26后，需要从数字中去除26代表的这个数
        if(tmp ==0) {
            tmp = 26;
            num-=1;
        }
        ch='A'+tmp-1;
        ret=ch+ret;
    }
    return ret;
}

void QPosition::set(int a,int b){
    row = a;
    col = b;
    cell = trans(col)+QString::number(row);
}
