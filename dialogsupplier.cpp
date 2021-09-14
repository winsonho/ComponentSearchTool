#include "dialogsupplier.h"
#include "ui_dialogsupplier.h"
#include <QDebug>

DialogSupplier::DialogSupplier(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::DialogSupplier)
{
    ui->setupUi(this);
    connect(ui->pushButton, SIGNAL(clicked()), this, SLOT(close()));
}

DialogSupplier::~DialogSupplier()
{
    delete ui;
}

void DialogSupplier::setSupplierList(QStringList supplierList)
{
    qDebug() << "setSupplierList";
    ui->tableWidget->setColumnCount(1);
    qDebug() << supplierList;
    for(int i=0; i<supplierList.count(); i++)
    {
        ui->tableWidget->insertRow(i);
        ui->tableWidget->setItem(i, 0, new QTableWidgetItem(supplierList.at(i)));

    }
    ui->tableWidget->resizeColumnsToContents();
}
