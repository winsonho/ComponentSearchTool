#include "mainwidget.h"
#include "ui_mainwidget.h"
#include <QFileDialog>
#include <QDebug>
#include <QMessageBox>
#include <QMenu>
#include <QRandomGenerator>
#include "dialogsupplier.h"
#include <QSet>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

#define MAJOR "1"
#define MINOR "06"
#define BUILD "0001"
#define SUPPLIER_LIST_FILE "supplier list.csv"
#define SUPPLIER_LIST_FIELD_NAME "Supplier Name"
#define SUPPLIER_FIELD_WIDTH    350

MainWidget::MainWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::MainWidget)
{
    ui->setupUi(this);
    QString version = QString("v%1.%2").arg(MAJOR).arg(MINOR) ;
    setWindowTitle("Component Search Tool:  " + version);

//    QColor color(QRandomGenerator::global()->bounded(255), QRandomGenerator::global()->bounded(255), QRandomGenerator::global()->bounded(255));
//qDebug() << color;
//    QBrush brush(color);
//qDebug() << brush << brush.color();

    // init
    m_txtImportFileName = ui->txtImportFileName;
    m_txtImportFileName1 = ui->txtImportFileName1;
    m_txtImportFileName2 = ui->txtImportFileName2;
    m_btnSelectFile = ui->btnSelectFile;
    m_btnSelectFile1 = ui->btnSelectFile1;
    m_btnSelectFile2 = ui->btnSelectFile2;
    m_txtSupplierImportFileName = ui->txtSupplierImportFileName;
    m_btnSupplierSelectFile = ui->btnSupplierSelectFile;
    m_btnSupplierImported = ui->btnSupplierImported;
    m_btnCheck = ui->btnCheck;
    m_btnImport = ui->btnImport;
    m_btnFixError = ui->btnFixError;
    m_btnExport = ui->btnExport;
    m_btnOneClick = ui->btnOneClickExport;
    m_btnSearch = ui->btnSearch;
    m_btnOutput = ui->btnOutput;

    m_twCSV = ui->twCSV;
    m_twCSV1 = ui->twCSV1;
    m_twCSV2 = ui->twCSV2;

    m_btnErrors[ErrorType::SPACE] = ui->btnErrorSpace;
    m_btnErrors[ErrorType::LOWERCASE] = ui->btnErrorLower;
    m_btnErrors[ErrorType::SPECIAL_CHAR] = ui->btnErrorChar;
    m_btnErrors[ErrorType::COMMA] = ui->btnErrorComma;
    m_btnErrors[ErrorType::CRITICAL_PART] = ui->btnErrorCriticalPart;
    m_btnErrors[ErrorType::SUPPLIER_LIST] = ui->btnErrorSupplier;

    m_chkErrors[ErrorType::SPACE] = ui->chkSpace;
    m_chkErrors[ErrorType::LOWERCASE] = ui->chkLower;
    m_chkErrors[ErrorType::SPECIAL_CHAR] = ui->chkChar;
    m_chkErrors[ErrorType::COMMA] = ui->chkComma;
    m_chkErrors[ErrorType::CRITICAL_PART] = ui->chkCriticalPart;
    m_chkErrors[ErrorType::SUPPLIER_LIST] = ui->chkSupplier;

    connect(m_btnSelectFile, SIGNAL(clicked()), this, SLOT(onSelectFileClicked()));
    connect(m_btnSelectFile1, SIGNAL(clicked()), this, SLOT(onSelectFileClicked()));
    connect(m_btnSelectFile2, SIGNAL(clicked()), this, SLOT(onSelectFileClicked()));
    connect(m_btnSupplierSelectFile, SIGNAL(clicked()), this, SLOT(onSupplierSelectFileClicked()));
    connect(m_btnSupplierImported, SIGNAL(clicked()), this, SLOT(onSupplierImportedClicked()));
    connect(m_btnCheck, SIGNAL(clicked()), this, SLOT(onCheckClicked()));
    connect(m_btnImport, SIGNAL(clicked()), this, SLOT(onImportClicked()));
    connect(m_btnFixError, SIGNAL(clicked()), this, SLOT(onFixErrorClicked()));
    connect(m_btnExport, SIGNAL(clicked()), this, SLOT(onExportClicked()));
    connect(m_btnOneClick, SIGNAL(clicked()), this, SLOT(onOneClickClicked()));
    connect(m_btnSearch, SIGNAL(clicked()), this, SLOT(onSearchClicked()));
    connect(m_btnOutput, SIGNAL(clicked()), this, SLOT(onExportClicked()));
    connect(m_twCSV, SIGNAL(cellDoubleClicked(int, int)), this, SLOT(cellSelected(int, int)));
    connect(m_twCSV, SIGNAL(currentCellChanged(int ,int, int, int)), this, SLOT(currentCellChanged(int, int, int, int)));


    for (int i=0; i<ErrorType::MAXCOUNT; i++)
    {
        connect(m_btnErrors[i], SIGNAL(clicked()), this, SLOT(onErrorsClicked()));
    }

    m_twCSV->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_twCSV, SIGNAL(customContextMenuRequested(QPoint)), SLOT(customMenuRequested(QPoint)));

    initError();

    QString supplierFileName = SUPPLIER_LIST_FILE;
    QDir dir = QCoreApplication::applicationDirPath();

    #ifdef Q_OS_MACOS
        dir.cdUp();
        dir.cdUp();
        dir.cdUp();
    #endif

    supplierFileName = dir.path() + QDir::separator() + supplierFileName;

    m_txtSupplierImportFileName->setText(supplierFileName);
    //loadSupplierList(supplierFileName);
}

void MainWidget::initError()
{
    for (int i=0; i<ErrorType::MAXCOUNT; i++)
    {
        errorInfos[i].errorColor = errorColors[i];
        errorInfos[i].errorType = (ErrorType)i;
//        errorInfos[i].errorCount = 0;
        errorInfos[i].errorString = errorStrings[i];
    }
}

void MainWidget::currentCellChanged(int currentRow, int currentColumn, int previousRow, int previousColumn)
{
//    qDebug() << currentRow << currentColumn;
//    qDebug() << previousRow << previousColumn;
    QWidget* widget = m_twCSV->cellWidget(previousRow, previousColumn);
    if ( widget != nullptr)
    {
        QLineEdit* edit = (QLineEdit *)widget;

//        qDebug() << "Found!! " << edit->text();
        QTableWidgetItem *newItem = new QTableWidgetItem(edit->text());

        m_twCSV->setItem(previousRow, previousColumn, newItem);
        m_twCSV->removeCellWidget(previousRow, previousColumn);
    }

}

void MainWidget::cellSelected(int row, int col)
{
//    qDebug() << row << col;
//    qDebug() << m_idxManufacturer << m_idxVendor;
    if (col == m_idxManufacturer || col == m_idxVendor)
    {
        qDebug() << m_idxManufacturer << m_idxVendor;
        QLineEdit *edit = new QLineEdit();
        QString str = m_twCSV->item(row, col)->text();
        edit->setText(str);
        edit->setCompleter(m_SupplierCompleter);
        m_twCSV->setCellWidget(row, col, edit);
    }

}
void MainWidget::onErrorsClicked()
{
    QObject* obj = sender();
    int idx = 0;
    for(int i=0; i<ErrorType::MAXCOUNT; i++)
    {
        if (obj == m_btnErrors[i])
        {
            idx = i;
            break;
        }
    }

//    qDebug() << idx ;

    if (m_ErrorFields[idx].count() == 0)
        return;

    if (m_idxErrors[idx] == m_ErrorFields[idx].count())
        m_idxErrors[idx] = 0;

//    qDebug() << m_idxErrors[idx];

    m_twCSV->scrollToItem(m_ErrorFields[idx].at(m_idxErrors[idx]));
    int col = m_twCSV->column(m_ErrorFields[idx].at(m_idxErrors[idx]));
    int row = m_twCSV->row(m_ErrorFields[idx].at(m_idxErrors[idx]));

//    qDebug() << row << col ;
    QModelIndex index = m_twCSV->model()->index(row, col);
    m_twCSV->selectionModel()->select(index, QItemSelectionModel::ClearAndSelect);

    ++m_idxErrors[idx];
}

void MainWidget::customMenuRequested(QPoint pos)
{
    QModelIndex index = m_twCSV->indexAt(pos);

    if (index.isValid())
    {
        QMenu *menu=new QMenu(this);
        menu->addAction("Delete", this, SLOT(onDelete()));
        menu->popup(m_twCSV->viewport()->mapToGlobal(pos));
    }
}

void MainWidget::onDelete()
{

    QModelIndexList selection = m_twCSV->selectionModel()->selectedRows();

    // Multiple rows can be selected
    for(int i=selection.count()-1; i>=0; i--)
    {
        QModelIndex index = selection.at(i);
//        qDebug() << index.row();
        m_twCSV->removeRow((index.row()));
    }
}

void MainWidget::onSelectFileClicked()
{
    //qDebug() << "here";
    QString filename = QFileDialog::getOpenFileName(this,
        tr("Open CSV file"), ".", tr("CSV Files (*.csv)"));

    if (sender()->objectName() == "btnSelectFile")
    {
        m_txtImportFileName->setText(filename);
    }
    else if (sender()->objectName() == "btnSelectFile1")
    {
        m_txtImportFileName1->setText(filename);
        onImportClicked(1);
    }
    else if (sender()->objectName() == "btnSelectFile2")
    {
        m_txtImportFileName2->setText(filename);
        onImportClicked(2);
    }
}

void MainWidget::onSupplierImportedClicked()
{
//    qDebug() << "onSupplierImportedClicked" ;
    if (m_SupplierList.count() > 0)
    {
        DialogSupplier dlg;
        dlg.setSupplierList(m_SupplierList);
        dlg.setModal(true);
        dlg.exec();
    }
}

void MainWidget::onSupplierSelectFileClicked()
{
    //qDebug() << "here";
    QString filename = QFileDialog::getOpenFileName(this,
        tr("Open Supplier List CSV file"), QDir::currentPath(), tr("CSV Files (*.csv)"));
    m_txtSupplierImportFileName->setText(filename);
    loadSupplierList(filename);
}

void MainWidget::onCheckClicked(bool isFix)
{
    bool bIsOK = true;
    if (m_twCSV->rowCount() == 0)
    {
        QMessageBox::information(NULL, "information", "請先匯入檔案!!!!", QMessageBox::Ok);
        return;
    }

    for (int i=0; i<ErrorType::MAXCOUNT ; i++)
    {
        m_ErrorFields[i].clear();
    }

    for (int x = 1; x < m_twCSV->rowCount(); x++)
    {
        // check item name
        int idx_ItemName = -1;
        if (m_CSVFieldMap.contains("item_name"))
            idx_ItemName = m_CSVFieldMap["item_name"];

        bIsOK &= itemNameCheck(x, idx_ItemName, isFix);

        // check vendor
        bIsOK &= supplierCheck(x, m_idxVendor, isFix);

        // check manufactor
        bIsOK &= supplierCheck(x, m_idxManufacturer, isFix);

        // check critical part
        int idx_CriticalPart = -1;
        if (m_CSVFieldMap.contains("critical part"))
            idx_CriticalPart = m_CSVFieldMap["critical part"];

        if (m_chkErrors[ErrorType::CRITICAL_PART]->isChecked() && m_twCSV->item(x, idx_CriticalPart) != NULL)
        {
            QString critical_part = m_twCSV->item(x, idx_CriticalPart)->text();

            if (isFix)
            {
                if (critical_part == "Y")
                    critical_part = "Yes";
                else if (critical_part == "N")
                    critical_part = "No";
                else if (!critical_part.isEmpty())
                    critical_part.clear();
                m_twCSV->item(x, idx_CriticalPart)->setText(critical_part);
                m_twCSV->item(x, idx_CriticalPart)->setBackground(normalColor);
            }

            if (critical_part != "Yes" && critical_part != "No" && !critical_part.isEmpty())
            {
                m_ErrorFields[ErrorType::CRITICAL_PART].append(m_twCSV->item(x, idx_CriticalPart));
                m_twCSV->item(x, idx_CriticalPart)->setBackground(errorColors[ErrorType::CRITICAL_PART]);
                bIsOK = false;
            }
        }
    }

    // update message button

    for (int i=0; i<ErrorType::MAXCOUNT ; i++)
    {
        m_btnErrors[i]->setText(QString(errorInfos[i].errorString).arg(m_ErrorFields[i].count()));
    }

    if (!isFix)
    {
        if (bIsOK)
            QMessageBox::information(NULL, "information", "檔案沒問題!!!!", QMessageBox::Ok);
        else
        {
            QMessageBox::critical(NULL, "information", "檔案格式有誤，請檢查有顏色的欄位!!!!", QMessageBox::Ok);
            // set focus
//            if (m_SpaceError.count() != 0)
//                onErrorSpaceClicked();
//            else if (m_LowerError.count() != 0)
//                onErrorLowerClicked();
//            else if (m_CharError.count() != 0)
//                onErrorCharClicked();
//            else if (m_CommaError.count() != 0)
//                onErrorCommaClicked();
//            else if (m_CriticalPartError.count() != 0)
//                onErrorCriticalPartClicked();
//            else if (m_ErrorSupplier.count() != 0)
//                onErrorSupplierClicked();
        }
    }
    else
    {
        QMessageBox::information(NULL, "information", "修正完成，請再執行檢查檔案一次!!!!", QMessageBox::Ok);
    }
}

bool MainWidget::supplierCheck(int row, int col, bool bIsFix)
{
    bool bIsOK= true;

    if (!m_chkErrors[ErrorType::SUPPLIER_LIST]->isChecked())
        return false;

    //qDebug() << row << col;
    if (m_twCSV->item(row, col) != NULL && !m_twCSV->item(row, col)->text().isEmpty())
    {
        QString name = m_twCSV->item(row, col)->text();
        //qDebug() << "[" << item_name << "]";// << "[" << item_name[item_name.length()-1] << "]";
        if (!m_SupplierList.contains(name))
        {
            if (bIsFix)
                m_twCSV->item(row, col)->setText(name.toUpper());
            else
                m_ErrorFields[ErrorType::SUPPLIER_LIST].append(m_twCSV->item(row, col));

            m_twCSV->item(row, col)->setBackground(errorInfos[ErrorType::SUPPLIER_LIST].errorColor);
            bIsOK = false;
        }
    }

    return bIsOK;
}

bool MainWidget::itemNameCheck(int row, int col, bool bIsFix)
{
    bool bIsOK= true;

    //qDebug() << row << col;
    if (m_twCSV->item(row, col) != NULL && !m_twCSV->item(row, col)->text().isEmpty())
    {
        QString item_name = m_twCSV->item(row, col)->text();
        //qDebug() << "[" << item_name << "]";// << "[" << item_name[item_name.length()-1] << "]";

        if (bIsFix)
        {
            // fix space
            if (m_chkErrors[ErrorType::SPACE]->isChecked())
                item_name = item_name.trimmed();

            // fix lower case
            if (m_chkErrors[ErrorType::LOWERCASE]->isChecked())
                item_name = item_name.toUpper();

            // fix contain ','
            if (m_chkErrors[ErrorType::COMMA]->isChecked())
                item_name = item_name.replace(',','/');

            if (m_chkErrors[ErrorType::SPECIAL_CHAR]->isChecked())
            {
                // fix '±'
                item_name = item_name.replace("±", "+-");
                // fix '°'
                item_name = item_name.replace("°", "degree");
            }

            m_twCSV->item(row, col)->setText(item_name);
            m_twCSV->item(row, col)->setBackground(normalColor);
        }

        // check leading and trail space for item name
        if (m_chkErrors[ErrorType::SPACE]->isChecked())
        {
            if (item_name[0].isSpace() || item_name[item_name.length()-1].isSpace())
            {
                m_ErrorFields[ErrorType::SPACE].append(m_twCSV->item(row, col));
                m_twCSV->item(row, col)->setBackground(errorInfos[ErrorType::SPACE].errorColor);
                bIsOK = false;
            }
        }

        // check lower case and special character.
        for (int i=0; i<item_name.length(); i++)
        {
            // check if contain lower case
            if (m_chkErrors[ErrorType::LOWERCASE]->isChecked() && item_name[i].isLower())
            {
                m_ErrorFields[ErrorType::LOWERCASE].append(m_twCSV->item(row, col));
                m_twCSV->item(row, col)->setBackground(errorInfos[ErrorType::LOWERCASE].errorColor);
                bIsOK = false;
                break;
            }

            int v_ascii = item_name[i].toLatin1();

            // check special character
            if (m_chkErrors[ErrorType::SPECIAL_CHAR]->isChecked())
            {
                if (v_ascii < 0 || v_ascii > 127)
                {
                    m_ErrorFields[ErrorType::SPECIAL_CHAR].append(m_twCSV->item(row, col));
                    m_twCSV->item(row, col)->setBackground(errorInfos[ErrorType::SPECIAL_CHAR].errorColor);
                    bIsOK = false;
                    break;
                }
            }
        }

        // check if contain ","
        if (m_chkErrors[ErrorType::SPECIAL_CHAR]->isChecked() && item_name.contains(QChar(',')))
        {
            m_ErrorFields[ErrorType::COMMA].append(m_twCSV->item(row, col));
            m_twCSV->item(row, col)->setBackground(errorInfos[ErrorType::COMMA].errorColor);
            bIsOK = false;
        }
    }

    return bIsOK;
}

QStringList MainWidget::parseCSV(const QString &string)
{
    enum State {Normal, Quote} state = Normal;
    QStringList fields;
    QString value;

    for (int i = 0; i < string.size(); i++)
    {
        const QChar current = string.at(i);

        // Normal state
        if (state == Normal)
        {
            // Comma
            if (current == ',')
            {
                // Save field
                fields.append(value.trimmed());
                value.clear();
            }

            // Double-quote
            else if (current == '"')
            {
                state = Quote;
                value += current;
            }

            // Other character
            else
                value += current;
        }

        // In-quote state
        else if (state == Quote)
        {
            // Another double-quote
            if (current == '"')
            {
                if (i < string.size())
                {
                    // A double double-quote?
                    if (i+1 < string.size() && string.at(i+1) == '"')
                    {
                        value += '"';

                        // Skip a second quote character in a row
                        i++;
                    }
                    else
                    {
                        state = Normal;
                        value += '"';
                    }
                }
            }

            // Other character
            else
                value += current;
        }
    }

    if (!value.isEmpty())
        fields.append(value.trimmed());

    // Quotes are left in until here; so when fields are trimmed, only whitespace outside of
    // quotes is removed.  The quotes are removed here.
    for (int i=0; i<fields.size(); ++i)
        if (fields[i].length()>=1 && fields[i].left(1)=='"')
        {
            fields[i]=fields[i].mid(1);
            if (fields[i].length()>=1 && fields[i].right(1)=='"')
                fields[i]=fields[i].left(fields[i].length()-1);
        }

    return fields;
}

bool MainWidget::readCSVRow (QTextStream &in, QStringList *row)
{

    static const int delta[][5] = {
        //  ,    "   \n    ?  eof
        {   1,   2,  -1,   0,  -1  }, // 0: parsing (store char)
        {   1,   2,  -1,   0,  -1  }, // 1: parsing (store column)
        {   3,   4,   3,   3,  -2  }, // 2: quote entered (no-op)
        {   3,   4,   3,   3,  -2  }, // 3: parsing inside quotes (store char)
        {   1,   3,  -1,   0,  -1  }, // 4: quote exited (no-op)
        // -1: end of row, store column, success
        // -2: eof inside quotes
    };

    row->clear();

    if (in.atEnd())
        return false;

    int state = 0, t;
    QChar ch;
    QString cell;

    while (state >= 0) {

        if (in.atEnd())
            t = 4;
        else {
            in >> ch;
            if (ch == ',') t = 0;
            else if (ch == '\"') t = 1;
            else if (ch == '\n') t = 2;
            else t = 3;
        }

        state = delta[state][t];

        switch (state) {
        case 0:
        case 3:
            cell += ch;
            break;
        case -1:
        case 1:
            row->append(cell);
            cell = "";
            break;
        }

    }

//    if (state == -2)
//        throw runtime_error("End-of-file found while inside quotes.");

    return true;

}

void MainWidget::onImportClicked(int index)
{
    QLineEdit* file;
    QTableWidget* csv;
    bool isComponent = false;

    if (index == 0) // import tool file
    {
        file =  m_txtImportFileName;
        csv = m_twCSV;
    }
    if (index == 1) // component tool file1
    {   
        m_Fields_CSV1.manufacturerItemNumber.clear();
        m_Fields_CSV1.vendorItemNumber.clear();
        m_Fields_CSV1.itemName.clear();
        m_Fields_CSV1.itemNumber.clear();
        file =  m_txtImportFileName1;
        csv = m_twCSV1;
        isComponent = true;
    }
    if (index == 2) // component tool file2
    {
        m_Fields_CSV2.manufacturerItemNumber.clear();
        m_Fields_CSV2.vendorItemNumber.clear();
        m_Fields_CSV2.itemName.clear();
        m_Fields_CSV2.itemNumber.clear();
        file =  m_txtImportFileName2;
        csv = m_twCSV2;
        isComponent = true;
    }

    if (file->text().isEmpty())
    {
        QMessageBox::information(NULL, "information", "請先選擇檔案!!!!", QMessageBox::Ok);
        return;
    }

    QFile importedCSV(file->text());
    QStringList rowData;
    csv->clear();
    csv->setRowCount(0);
    rowData.clear();

    importedCSV.open(QFile::ReadOnly);

    QTextStream in(&importedCSV);
    int row = 0;

    while (readCSVRow(in, &rowData))
    {
        csv->setColumnCount(rowData.size());
        csv->insertRow(row);
        for (int y = 0; y < rowData.size(); y++)
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(rowData[y]);

            if (row == 0)
            {
                if (index == 0) // inport tool
                    m_CSVFieldMap.insert(rowData[y], y);
                else if (index == 1) // CSV 1
                {
                    m_CSV1FieldMap.insert(rowData[y], y);
                    if (rowData[y].trimmed() == "manufacturer_item_number")
                    {
                        m_Fields_CSV1.idxManufacturerItemNumber = y;
                    }
                    else if (rowData[y].trimmed() == "vendor_item_number")
                    {
                        m_Fields_CSV1.idxVendorItemNumber = y;
                    }
                    else if (rowData[y].trimmed() == "item_number")
                    {
                        m_Fields_CSV1.idxItemNumber = y;
                    }
                    else if (rowData[y].trimmed() == "item_name")
                    {
                        m_Fields_CSV1.idxItemName = y;
                    }
                }
                else if (index == 2) // CSV 2
                {
                    m_CSV2FieldMap.insert(rowData[y], y);
                    if (rowData[y].trimmed() == "manufacturer_item_number")
                    {
                        m_Fields_CSV2.idxManufacturerItemNumber = y;
                    }
                    else if (rowData[y].trimmed() == "vendor_item_number")
                    {
                        m_Fields_CSV2.idxVendorItemNumber = y;
                    }
                    else if (rowData[y].trimmed() == "item_number")
                    {
                        m_Fields_CSV2.idxItemNumber = y;
                    }
                    else if (rowData[y].trimmed() == "item_name")
                    {
                        m_Fields_CSV2.idxItemName = y;
                    }
                }
            }
            else
            {
                // append data to CSV1
                if (index == 1)
                {
                    if (y == m_Fields_CSV1.idxManufacturerItemNumber)
                        m_Fields_CSV1.manufacturerItemNumber.append(rowData[y].trimmed());
                    else if (y == m_Fields_CSV1.idxVendorItemNumber)
                        m_Fields_CSV1.vendorItemNumber.append(rowData[y].trimmed());
                    else if (y == m_Fields_CSV1.idxItemNumber)
                        m_Fields_CSV1.itemNumber.append(rowData[y].trimmed());
                    else if (y == m_Fields_CSV1.idxItemName)
                        m_Fields_CSV1.itemName.append(rowData[y].trimmed());
                }
                else if (index == 2) // append data to CSV2
                {
                    if (y == m_Fields_CSV2.idxManufacturerItemNumber)
                        m_Fields_CSV2.manufacturerItemNumber.append(rowData[y].trimmed());
                    else if (y == m_Fields_CSV2.idxVendorItemNumber)
                        m_Fields_CSV2.vendorItemNumber.append(rowData[y].trimmed());
                    else if (y == m_Fields_CSV2.idxItemNumber)
                        m_Fields_CSV2.itemNumber.append(rowData[y].trimmed());
                    else if (y == m_Fields_CSV2.idxItemName)
                        m_Fields_CSV2.itemName.append(rowData[y].trimmed());
                }
            }
            newItem->setForeground(QBrush(Qt::black));
            newItem->setBackground(QBrush(Qt::white));
            csv->setItem(row, y, newItem);
//csv->item(row, y)->setForeground(QBrush(Qt::white));
//csv->item(row, y)->setBackground(QBrush(Qt::black));
//            qDebug() << csv->item(row, y)->foreground();
//            qDebug() << csv->item(row, y)->background();
        }

        row++;
    }
    importedCSV.close();
    csv->resizeColumnsToContents();

//    qDebug() << m_ManuItemNumber1;
//    qDebug() << m_ManuItemNumber2;

    if (!isComponent)
    {
        // vendor index
        m_idxVendor = -1;
        if (m_CSVFieldMap.contains("vendor"))
        {
            m_idxVendor = m_CSVFieldMap["vendor"];
            csv->setColumnWidth(m_idxVendor, SUPPLIER_FIELD_WIDTH);
        }
        // manufacturer index
        m_idxManufacturer = -1;
        if (m_CSVFieldMap.contains("manufacturer"))
        {
            m_idxManufacturer = m_CSVFieldMap["manufacturer"];
            csv->setColumnWidth(m_idxManufacturer, SUPPLIER_FIELD_WIDTH);
        }
    }
//    qDebug() << "m_idxManuItemNumber1: " << m_idxManuItemNumber1;

    if (index == 1)
    {
        m_twCSV1->scrollToItem(m_twCSV1->item(0, m_Fields_CSV1.idxManufacturerItemNumber), QAbstractItemView::PositionAtCenter);
        m_twCSV1->selectionModel()->select(m_twCSV1->model()->index(0, m_Fields_CSV1.idxManufacturerItemNumber), QItemSelectionModel::ClearAndSelect);
        csv->resizeColumnsToContents();
        //m_twCSV1->scrollTo(m_twCSV1->model()->index(1, m_idxManuItemNumber1));
    }
    if (index == 2)
    {
        m_twCSV2->scrollToItem(m_twCSV2->item(0, m_Fields_CSV2.idxManufacturerItemNumber), QAbstractItemView::PositionAtCenter);
        m_twCSV2->selectionModel()->select(m_twCSV2->model()->index(0, m_Fields_CSV2.idxManufacturerItemNumber), QItemSelectionModel::ClearAndSelect);
        csv->resizeColumnsToContents();
        //m_twCSV1->scrollTo(m_twCSV1->model()->index(1, m_idxManuItemNumber1));
    }
}

void MainWidget::onFixErrorClicked()
{
    onCheckClicked(true);
}

void MainWidget::onExportClicked(bool isExcel)
{
    QTableWidget *tw = nullptr;
    QString defaultFile;

    if (sender()->objectName() == "btnExport")
    {
        tw = m_twCSV;
        defaultFile = m_txtImportFileName->text();
    }
    else if (sender()->objectName() == "btnOutput")
    {
        tw = m_twCSV2;
        defaultFile = m_txtImportFileName2->text();
    }
    else
    {
        qDebug() << "Don't know who sender: " << sender()->objectName() << " is!";
        return;
    }

    if (tw->rowCount() == 0)
    {
        QMessageBox::information(NULL, "information", "請先匯入檔案!!!!", QMessageBox::Ok);
        return;
    }

    QDir dir(defaultFile);
    QString filename = QFileDialog::getSaveFileName(this,
        tr("Save file"), dir.absoluteFilePath(defaultFile), tr("CSV Files (*.csv);;Excel Files (*.xlsx)"));

    if (!filename.isEmpty())
    {
        // csv file
        if (filename.endsWith(".csv"))
        {
            QFile exportCSV(filename);

            if (exportCSV.open(QIODevice::Append))
            {
                QTextStream stream(&exportCSV);
                for (int i=0; i<tw->rowCount(); i++)
                {
                    QString data;
                    int j=0;
                    for(j=0; j<tw->columnCount()-1; j++)
                    {
                        qDebug() << i << j << tw->item(i,j)->text();
                        QString str = tw->item(i,j)->text();
                        str.replace('"', "\"\"");
                        if (str.contains(',') || str.contains('\n'))
                            str = "\"" + str + "\"";
                       data += str + ',';
                    }
                    // last colummn
                    QString str = tw->item(i,j)->text();
                    str.replace('"', "\"\"");
                    if (str.contains(',') || str.contains('\n'))
                        str = "\"" + str + "\"";
                    data += str;

                    qDebug() << data;
                    stream << data << "\n";
                }
                exportCSV.close();
            }
        }
        // excel file
        else if (filename.endsWith(".xlsx"))
        {
            Document doc;
            for (int i=0; i<tw->rowCount(); i++)
            {
                QString data;
                for(int j=0; j<tw->columnCount(); j++)
                {
                    QTableWidgetItem *item = tw->item(i,j);
                    Format format;
                    format.setFontColor(item->foreground().color());
                    format.setPatternBackgroundColor(item->background().color());
                    format.setVerticalAlignment(Format::AlignVCenter);
                    QString str = item->text();
                    doc.write(i+1, j+1, QVariant(str), format);
                }
            }
            doc.autosizeColumnWidth();
            doc.saveAs(filename);
        }
        QMessageBox::information(NULL, "information", "Export to " + filename + " Done!!!!", QMessageBox::Ok);
    }
}

void MainWidget::onOneClickClicked()
{
    onSelectFileClicked();
    onImportClicked();
    onCheckClicked(true);
    onExportClicked();
    //QMessageBox::information(NULL, "information", "求我啊，哈哈哈!!!!", QMessageBox::Ok);
}


void MainWidget::loadSupplierList(QString supplierFileName)
{
    QFile file(supplierFileName);

    if (!file.exists())
    {
        QMessageBox::information(NULL, "information", "Supplier list檔案: " + file.fileName() + " 不存在!!!!", QMessageBox::Ok);
        return;
    }

    m_SupplierList.clear();

    QFile importedCSV(supplierFileName);
    QStringList rowData;
    rowData.clear();

    importedCSV.open(QFile::ReadOnly);

    QTextStream in(&importedCSV);
    int row = 0;
    int idxSupplierName = -1;
    while (readCSVRow(in, &rowData))
    {
        // get first row to check "Supplier Name"
        if (row == 0)
        {
            for (int y = 0; y < rowData.size(); y++)
            {
                if (rowData[y].compare(SUPPLIER_LIST_FIELD_NAME) == 0)
                {
                    idxSupplierName = y;
                }
            }
        }
        else
        {
            if (idxSupplierName != -1)
                m_SupplierList.append(rowData[idxSupplierName]);
        }
        row++;
    }
    importedCSV.close();
//    qDebug() << m_SupplierList;
    if (m_SupplierCompleter != nullptr)
        delete m_SupplierCompleter;
    m_SupplierCompleter = new QCompleter(m_SupplierList, this);
    m_btnSupplierImported->setText(QString("已匯入Supplier(%1)筆").arg(m_SupplierList.count()));
//    ui->lineEdit->setCompleter(m_SupplierCompleter);
}

void MainWidget::checkDuplicate()
{
//    QMap<QString,int> countOfManuItemStrings;
    QMap<QString,QSet<QString>> vendorItemStringsSet;
    QMap<QString,QSet<QString>> itemStringsSet;
    QMap<QString,QSet<QString>> nameStringsSet;

    //QMap<QString,int> countOfItemStrings;
    for (int i=0; i<m_Fields_CSV1.manufacturerItemNumber.count(); i++)
    {
        if (!m_Fields_CSV1.vendorItemNumber.at(i).trimmed().isEmpty())
            vendorItemStringsSet[m_Fields_CSV1.manufacturerItemNumber.at(i)].insert(m_Fields_CSV1.vendorItemNumber.at(i));
        if (!m_Fields_CSV1.itemNumber.at(i).trimmed().isEmpty())
            itemStringsSet[m_Fields_CSV1.manufacturerItemNumber.at(i)].insert(m_Fields_CSV1.itemNumber.at(i));
        if (!m_Fields_CSV1.itemName.at(i).trimmed().isEmpty())
            nameStringsSet[m_Fields_CSV1.manufacturerItemNumber.at(i)].insert(m_Fields_CSV1.itemName.at(i));
    }
//    qDebug() << vendorItemStringsSet;
//    qDebug() << itemStringsSet;

    QMap<QString, QSet<QString>>::const_iterator v = vendorItemStringsSet.constBegin();

    while(v != vendorItemStringsSet.constEnd())
    {
        if (v.value().count() >1)
        {
            foreach (const QString &str, v.value())
            {
                m_VendorItemNumberMap[v.key()].append(str);
            }
        }
        ++v;
    }

    QMap<QString, QSet<QString>>::const_iterator i = itemStringsSet.constBegin();

    while(i != itemStringsSet.constEnd())
    {
        if (i.value().count() >1)
        {
            foreach (const QString &str, i.value())
            {
                m_ItemNumberMap[i.key()].append(str);
            }
        }
        ++i;
    }

    QMap<QString, QSet<QString>>::const_iterator n = nameStringsSet.constBegin();

    while(n != nameStringsSet.constEnd())
    {
        if (n.value().count() >1)
        {
            foreach (const QString &str, n.value())
            {
                m_ItemNameMap[n.key()].append(str);
            }
        }
        ++n;
    }
    qDebug() << m_VendorItemNumberMap;
    qDebug() << m_ItemNumberMap;
    qDebug() << m_ItemNameMap;
}

void MainWidget::onSearchClicked()
{
    // reset data first
    onImportClicked(1);
    onImportClicked(2);

    if (m_twCSV1->rowCount() == 0 || m_twCSV2->rowCount() == 0)
    {
        QMessageBox::information(NULL, "information", "請先匯入檔案!!!!", QMessageBox::Ok);
        return;
    }

    if (m_Fields_CSV1.manufacturerItemNumber.count() == 0 || m_Fields_CSV2.manufacturerItemNumber.count() == 0)
    {
        QMessageBox::information(NULL, "information", "檔案沒有 manufacturer_item_number 欄位!!!!", QMessageBox::Ok);
        return;
    }

    if (m_Fields_CSV1.itemNumber.count() == 0 || m_Fields_CSV2.itemNumber.count() == 0)
    {
        QMessageBox::information(NULL, "information", "檔案沒有 itemNumber 欄位!!!!", QMessageBox::Ok);
        return;
    }

    if (m_Fields_CSV1.itemName.count() == 0 || m_Fields_CSV2.itemName.count() == 0)
    {
        QMessageBox::information(NULL, "information", "檔案沒有 itemName 欄位!!!!", QMessageBox::Ok);
        return;
    }

    if (m_Fields_CSV1.vendorItemNumber.count() == 0 || m_Fields_CSV2.vendorItemNumber.count() == 0)
    {
        QMessageBox::information(NULL, "information", "檔案沒有 vendorItemNumber 欄位!!!!", QMessageBox::Ok);
        return;
    }

    QStringList list = m_Fields_CSV1.manufacturerItemNumber;

    checkDuplicate();
    list.removeDuplicates();
    // remove "none" && space && "N/A"
    list.removeAll("");
    list.removeAll("none");
    list.removeAll("N/A");

    int manuItemCount = 0;
    int manuItemTotalCount = 0;
    int vendorItemCount = 0;
    int itemNumberCount = 0;
    int itemNameCount = 0;

    for (int i = 0; i < list.count(); i++)
    {
        // search CVS 2
        //int idx = -1;
        bool bFound = false;
        QColor color(QRandomGenerator::global()->bounded(128, 255), QRandomGenerator::global()->bounded(128, 255), QRandomGenerator::global()->bounded(128, 255));
        //qDebug() << "Search: " << list.at(i).trimmed();

        for (int idx = 0 ; idx < m_Fields_CSV2.manufacturerItemNumber.size(); idx++)
        {
            idx = m_Fields_CSV2.manufacturerItemNumber.indexOf(list.at(i).trimmed(), idx);

//            qDebug() << "m_ManuItemNumber2: " << idx ;
            if (idx < 0)
                break ;

            int idx1 = m_Fields_CSV1.manufacturerItemNumber.indexOf(list.at(i).trimmed());
            m_twCSV2->item(idx+1, m_Fields_CSV2.idxManufacturerItemNumber)->setBackground(QBrush(color));
            // also set item number
            m_twCSV2->item(idx+1, m_Fields_CSV2.idxItemNumber)->setText(m_Fields_CSV1.itemNumber.at(idx1));
            //m_twCSV2->item(idx+1, m_Fields_CSV2.idxItemName)->setText(m_Fields_CSV1.itemName.at(idx));

            // check item_name and vendor_item_number
            QString str;

            if (m_VendorItemNumberMap.contains(list.at(i)))
            {
                str = m_VendorItemNumberMap[list.at(i)].at(0);
                for (int idxVendor=1; idxVendor<m_VendorItemNumberMap[list.at(i)].count(); idxVendor++ )
                {
                    str += "\r\n" + m_VendorItemNumberMap[list.at(i)].at(idxVendor);
                }
                if (!m_VendorItemNumberMap[list.at(i)].contains(m_Fields_CSV2.vendorItemNumber.at(idx).trimmed()))
                    str += "\r\n" + m_Fields_CSV2.vendorItemNumber.at(idx).trimmed();

                m_twCSV2->item(idx+1, m_Fields_CSV2.idxVendorItemNumber)->setText(str);
                m_twCSV2->item(idx+1, m_Fields_CSV2.idxVendorItemNumber)->setBackground(QBrush(Qt::red));

                vendorItemCount++;
            }

            if (m_ItemNumberMap.contains(list.at(i)))
            {
                str = m_ItemNumberMap[list.at(i)].at(0);
                for (int idxItem=1; idxItem<m_ItemNumberMap[list.at(i)].count(); idxItem++ )
                {
                    str += "\r\n" + m_ItemNumberMap[list.at(i)].at(idxItem);
                }
                if (!m_ItemNumberMap[list.at(i)].contains(m_Fields_CSV2.itemNumber.at(idx).trimmed()))
                    str += "\r\n" + m_Fields_CSV2.itemNumber.at(idx).trimmed();

                m_twCSV2->item(idx+1, m_Fields_CSV2.idxItemNumber)->setText(str);
                m_twCSV2->item(idx+1, m_Fields_CSV2.idxItemNumber)->setBackground(QBrush(Qt::red));
                itemNumberCount++;
            }

            if (m_ItemNameMap.contains(list.at(i)))
            {
                str = m_ItemNameMap[list.at(i)].at(0);
                for (int idxName=1; idxName<m_ItemNameMap[list.at(i)].count(); idxName++ )
                {
                    str += "\r\n" + m_ItemNameMap[list.at(i)].at(idxName);
                }
                if (!m_ItemNameMap[list.at(i)].contains(m_Fields_CSV2.itemName.at(idx).trimmed()))
                    str += "\r\n" + m_Fields_CSV2.itemName.at(idx).trimmed();

                m_twCSV2->item(idx+1, m_Fields_CSV2.idxItemName)->setText(str);
                m_twCSV2->item(idx+1, m_Fields_CSV2.idxItemName)->setBackground(QBrush(Qt::red));
                itemNameCount++;
            }

            bFound = true;
            manuItemTotalCount++;
        }

        if (bFound)
        {
            for (int idx = 0 ; idx < m_Fields_CSV1.manufacturerItemNumber.size(); idx++)
            {
                idx = m_Fields_CSV1.manufacturerItemNumber.indexOf(list.at(i).trimmed(), idx);
//                qDebug() << "m_ManuItemNumber1: " << idx ;

                if (idx < 0)
                    break ;

                if (m_VendorItemNumberMap.contains(list.at(i).trimmed()))
                    m_twCSV1->item(idx+1, m_Fields_CSV1.idxVendorItemNumber)->setBackground(QBrush(color));

                if (m_ItemNumberMap.contains(list.at(i).trimmed()))
                    m_twCSV1->item(idx+1, m_Fields_CSV1.idxItemNumber)->setBackground(QBrush(color));

                if (m_ItemNameMap.contains(list.at(i).trimmed()))
                    m_twCSV1->item(idx+1, m_Fields_CSV1.idxItemName)->setBackground(QBrush(color));

                m_twCSV1->item(idx+1, m_Fields_CSV1.idxManufacturerItemNumber)->setBackground(QBrush(color));
            }
            m_VendorItemNumberMap.remove(list.at(i).trimmed());
            m_ItemNumberMap.remove(list.at(i).trimmed());
            m_ItemNameMap.remove(list.at(i).trimmed());
            manuItemCount++;
        }
    }

    // check the original errors.
    int orgItemNumberCount = 0;
    int orgVendorItemNumberCount = 0;
    int orgItemNameCount = 0;
    for (int idx = 0 ; idx < m_Fields_CSV1.manufacturerItemNumber.size(); idx++)
    {
        if (m_VendorItemNumberMap.contains(m_Fields_CSV1.manufacturerItemNumber.at(idx)))
        {
            m_twCSV1->item(idx+1, m_Fields_CSV1.idxVendorItemNumber)->setBackground(QBrush(Qt::red));
            ++orgVendorItemNumberCount;
        }
        if (m_ItemNumberMap.contains(m_Fields_CSV1.manufacturerItemNumber.at(idx)))
        {
            m_twCSV1->item(idx+1, m_Fields_CSV1.idxItemNumber)->setBackground(QBrush(Qt::red));
            ++orgItemNumberCount;
        }
        if (m_ItemNameMap.contains(m_Fields_CSV1.manufacturerItemNumber.at(idx)))
        {
            m_twCSV1->item(idx+1, m_Fields_CSV1.idxItemName)->setBackground(QBrush(Qt::red));
            ++orgItemNameCount;
        }

    }


    m_twCSV1->resizeRowsToContents();
    m_twCSV2->resizeRowsToContents();

    QString msg;
    if (manuItemCount > 0)
    {
        msg = QString("總共有: %1 個match, 總個數為: %2 個item").arg(manuItemCount).arg(manuItemTotalCount);
    }
    else
    {
        msg = "找不到相同的manufacturer_item_number";
    }
    QMessageBox::information(NULL, "information", msg, QMessageBox::Ok);

    if (vendorItemCount > 0 || itemNameCount > 0 || itemNumberCount > 0)
    {
        msg = QString("慘了慘了!!\n"
                        "總共 %1 個 item_number 不一樣\n"
                        "總共 %2 個 item_name 不一樣\n"
                        "總共 %3 個 vendor_item_name 不一樣").arg(itemNumberCount).arg(itemNameCount).arg(vendorItemCount);
        QMessageBox::critical(NULL, "Error", msg, QMessageBox::Ok);
    }

    if (orgItemNumberCount > 0 || orgVendorItemNumberCount > 0 || orgItemNameCount > 0)
    {
        msg = QString("非常糟糕，慘了慘了，原始資料發現錯誤了!!\n"
                        "總共 %1 個 item_number 不一樣\n"
                        "總共 %2 個 item_name 不一樣\n"
                        "總共 %3 個 vendor_item_name 不一樣").arg(orgItemNumberCount).arg(orgItemNameCount).arg(orgVendorItemNumberCount);
        QMessageBox::critical(NULL, "Error", msg, QMessageBox::Ok);
    }

}

MainWidget::~MainWidget()
{
    delete ui;
}
