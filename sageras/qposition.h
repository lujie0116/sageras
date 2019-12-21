#ifndef QPOSITION_H
#define QPOSITION_H

#include <QString>

class QPosition{
public:
    QPosition(int a=1,int b=1);
    QString trans(int num);
    void set(int a,int b);

    int row = 1;
    int col = 1;
    QString cell = trans(col)+QString::number(row);
};

#endif // QPOSITION_H
