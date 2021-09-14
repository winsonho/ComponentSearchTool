#ifndef DIALOGSUPPLIER_H
#define DIALOGSUPPLIER_H

#include <QDialog>

namespace Ui {
class DialogSupplier;
}

class DialogSupplier : public QDialog
{
    Q_OBJECT

public:
    explicit DialogSupplier(QWidget *parent = nullptr);
    ~DialogSupplier();
    void setSupplierList(QStringList supplierList);

private:
    Ui::DialogSupplier *ui;
    //QStringList m_SupplierList;
};

#endif // DIALOGSUPPLIER_H
