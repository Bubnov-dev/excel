#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QMessageBox>
#include <ActiveQt\QAxObject>
#include <QDir>
#include<QDebug>
#include<QTableWidgetItem>
#include<QInputDialog>
#include <QFileDialog>
#include <QGridLayout>
#include <QShortcut>


QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void buttoms();

    void buttoms2();

    void addToExcel(QString list,int row,int column,QTableWidgetItem* value);

    void removeFromExcel(QString list,char *comand,int number);

    void removeFromExcelCells(QString list,int row,int column);

    void key();
public slots:

    void saveAs();

    void save();

    void addRow();

    void removeRow();

    void removeColumn();

    void addColumn();

    void writeCell();

    void renameTab();

    void addTab();

    void removeTab();

    void changeTable();

    void deleteItems();

    void questions();

private:
    Ui::MainWindow *ui;

    QUrl dir;
    QShortcut *del;

    QInputDialog inputDialog;
    QTableWidget *table;
    QTabWidget *tabWidget;

    QAxObject *mExcel;
    QAxObject *workbooks;
    QAxObject *workbook;
    QAxObject *mSheets;
    QAxObject *workSheet;
    QSet<QString> set;
};
#endif // MAINWINDOW_H
