#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QtSql/QSqlDatabase>
#include <QtSql/QSqlDriver>

#include <QTreeWidgetItem>

#include <QStandardItemModel>  //实现通用的二维数据的管理功能。
#include <QtSql/QSqlTableModel>

#include <QAxObject>

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow( QWidget* parent = nullptr );
    ~MainWindow();

private slots:
    void initTableview();
    void treeWidgetClicked( QTreeWidgetItem* item );
    bool distinguishNodes( QString filtername );
    void fatherNodeClicked( QTreeWidgetItem* item, QString filtername );
    void childrenNodeClicked( QTreeWidgetItem* item, QString filtername );

    void creatNewTable( QAxObject* work_sheet, QString filename );
    void InsertData( QAxObject* work_sheet, QString filename );
    bool isTableExists( QString& table );

    void on_pushButton_import_clicked();


private:
    Ui::MainWindow* ui;

    QSqlDatabase dbMYSQL;

    int                 currentcolumn = 0;
    QStandardItemModel* model;
    QSqlTableModel*     mymodel;
    int                 tablecount, canshucount;
};
#endif  // MAINWINDOW_H
