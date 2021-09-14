#ifndef MAINWIDGET_H
#define MAINWIDGET_H

#include <QWidget>
#include <QLineEdit>
#include <QPushButton>
#include <QCheckBox>
#include <QTableWidget>
#include <QMap>
#include <QCompleter>
#include "Common.h"

struct ComponentFields{
    int idxItemNumber = -1;
    QStringList itemNumber;
    int idxItemName = -1;
    QStringList itemName;
    int idxVendorItemNumber = -1;
    QStringList vendorItemNumber;
    int idxManufacturerItemNumber = -1;
    QStringList manufacturerItemNumber;
};

namespace Ui {
class MainWidget;
}

class MainWidget : public QWidget
{
    Q_OBJECT

public:
    explicit MainWidget(QWidget *parent = nullptr);
    ~MainWidget();

private slots:
    void onSelectFileClicked();
    void onSupplierSelectFileClicked();
    void onSupplierImportedClicked();
    void onCheckClicked(bool isFix=false);
    void onImportClicked(int index=0);
    void onExportClicked(bool isExcel=false);
    void onOneClickClicked();
    void onFixErrorClicked();
    void customMenuRequested(QPoint pos);
    void onDelete();
    void onErrorsClicked();
    void cellSelected(int row, int col);
    void currentCellChanged(int currentRow, int currentColumn, int previousRow, int previousColumn);
    void onSearchClicked();

private:
    Ui::MainWidget *ui;
    QLineEdit* m_txtImportFileName;
    QLineEdit* m_txtImportFileName1;
    QLineEdit* m_txtImportFileName2;
    QLineEdit* m_txtSupplierImportFileName;
    QPushButton* m_btnSelectFile;
    QPushButton* m_btnSelectFile1;
    QPushButton* m_btnSelectFile2;
    QPushButton* m_btnSupplierSelectFile;
    QPushButton* m_btnSupplierImported;
    QPushButton* m_btnImport;
    QPushButton* m_btnCheck;
    QPushButton* m_btnFixError;
    QPushButton* m_btnExport;
    QPushButton* m_btnOneClick;
    QPushButton* m_btnSearch;
    QPushButton* m_btnOutput;
    QTableWidget* m_twCSV;
    QTableWidget* m_twCSV1;
    QTableWidget* m_twCSV2;
    QMap<QString, int> m_CSVFieldMap;
    QMap<QString, int> m_CSV1FieldMap;
    QMap<QString, int> m_CSV2FieldMap;
//    QStringList m_ManuItemNumber1;
//    QStringList m_ManuItemNumber2;
//    QStringList m_ManuItemNumber;
//    int m_idxManuItemNumber1 = -1;
//    int m_idxManuItemNumber2 = -1;

    ComponentFields m_Fields_CSV1, m_Fields_CSV2;
    QStringList parseCSV(const QString &string);

    QList<QTableWidgetItem*> m_ErrorFields[ErrorType::MAXCOUNT];
    QPushButton* m_btnErrors[ErrorType::MAXCOUNT];
    int m_idxErrors[ErrorType::MAXCOUNT] = {0} ;

    bool itemNameCheck(int row, int col, bool bIsFix=false);
    bool readCSVRow (QTextStream &in, QStringList *row);
    QStringList m_SupplierList;
    void loadSupplierList(QString supplierFileName);
    bool supplierCheck(int row, int col, bool bIsFix=false);
    QCompleter* m_SupplierCompleter = nullptr;

    void initError();
    QCheckBox* m_chkErrors[ErrorType::MAXCOUNT];
    QList<QTableWidgetItem*> m_ErrorWidgetItems[ErrorType::MAXCOUNT];
    ErrorInfo errorInfos[ErrorType::MAXCOUNT];
    QBrush errorColors[ErrorType::MAXCOUNT] = {
        QBrush(QColor(255, 0, 0)),
        QBrush(QColor(0, 255, 0)),
        QBrush(QColor(0, 0, 255)),
        QBrush(QColor(255, 255, 0)),
        QBrush(QColor(0, 255, 255)),
        QBrush(QColor(255, 0, 255)),
    };
    QString errorStrings[ErrorType::MAXCOUNT] = {
        QString("包含空白(%1)"),
        QString("包含小寫(%1)"),
        QString("包含特殊字元(%1)"),
        QString("包含逗號(%1)"),
        QString("Critical Part(%1)"),
        QString("Supplier List(%1)"),
    };
    QBrush normalColor = QBrush(QColor(255, 255, 255));
    int m_idxVendor = -1;
    int m_idxManufacturer = -1;
    void checkDuplicate();
    QMap<QString, QStringList> m_VendorItemNumberMap;
    QMap<QString, QStringList> m_ItemNumberMap;
    QMap<QString, QStringList> m_ItemNameMap;
};

#endif // MAINWIDGET_H
