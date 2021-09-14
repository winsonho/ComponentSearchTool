#ifndef COMMON_H
#define COMMON_H
#include <QBrush>

enum ErrorType{
    SPACE,
    LOWERCASE,
    SPECIAL_CHAR,
    COMMA,
    CRITICAL_PART,
    SUPPLIER_LIST,
    MAXCOUNT,
};

struct ErrorInfo{
    ErrorType errorType;
    QBrush errorColor;
    QString errorString;
//    int errorCount;
};

#endif // COMMON_H
