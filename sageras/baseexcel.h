#ifndef BASEEXCEL_H
#define BASEEXCEL_H

#include <ActiveQt/QAxObject>

class BaseExcel
{
public:
    BaseExcel();
protected:

public:
    QAxObject*  excel;
    QAxObject*  books;
    QAxObject*  book;
    QAxObject*  sheets;
    QAxObject*  sheet;
};

#endif // BASEEXCEL_H
