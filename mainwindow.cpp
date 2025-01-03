#include "mainwindow.h"
#include "./ui_mainwindow.h"

#include <QtSql/QSqlDatabase>
#include <QtSql/QSqlDriver>
#include <QSqlField>
#include <QtDebug>
#include <QSqlError>
#include <QSqlQuery>
#include <QFile>

#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxchartsheet.h"
#include "xlsxdocument.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

#include <QTreeWidget>

#include <QStandardItemModel>  //实现通用的二维数据的管理功能。
#include <QSqlTableModel>
#include <QSqlRecord>

#include <QFileDialog>
#include <QMessageBox>


// #define chosedtable;
QStringList tablelist, canshulist;
bool        parentNodeSelected = false;

MainWindow::MainWindow( QWidget* parent ) : QMainWindow( parent ), ui( new Ui::MainWindow ) {
    ui->setupUi( this );

    initTableview();  // 初始化默认状态
    ui->treeWidget->setHeaderLabel( "表格名" );
    // MySQL
    dbMYSQL = QSqlDatabase::addDatabase( "QMYSQL" );  // 创建数据库对象
    dbMYSQL.setHostName( "localhost" );               // 为本机的 IP
    dbMYSQL.setPort( 3306 );                          // 端口号，一般数据库都为 3306
    dbMYSQL.setDatabaseName( "mydatabase" );          // 自己设的数据库名
    dbMYSQL.setUserName( "root" );                    // 登录用户名 在创建数据库时设置的用户名和密码
    dbMYSQL.setPassword( "Cz1253709179." );           // 登录密码
    if ( dbMYSQL.open() ) {
        qDebug() << "open successful";  // 数据库连接成功
    }
    else {
        qDebug() << "error" << dbMYSQL.lastError().text();  // 连接失败打印error信息
    }

    // 【信号与槽连接】
    QObject::connect( ui->treeWidget, &QTreeWidget::itemClicked, this, &MainWindow::treeWidgetClicked );

    mymodel = new QSqlTableModel( this, dbMYSQL );
    mymodel->setTable( "shanmuban15" );
    mymodel->setEditStrategy( QSqlTableModel::OnManualSubmit );
    mymodel->select();
    model = new QStandardItemModel();  // QStandardItemModel 是包含单元格的容器（在这里可以看作表）
}

void MainWindow::initTableview() {
    QStandardItemModel* model = new QStandardItemModel( ui->tableView );
    // model->setHorizontalHeaderLabels(QStringList()<<QStringLiteral("dgs")<<QStringLiteral("info"));
    // model->setHorizontalHeaderLabels(QStringList()<<QStringLiteral(" "));
    QStandardItem*      first = new QStandardItem( "" );  // 添加
    model->appendRow( first );
    ui->tableView->setEditTriggers( QAbstractItemView::NoEditTriggers );  // 设置为不可修改
    ui->tableView->setModel( model );
}

// 接收 itemClicked() 信号函数传递过来的 item 参数
void MainWindow::treeWidgetClicked( QTreeWidgetItem* item ) {
    QString filtername = item->text( 0 ).trimmed();  // 节点名，子节点时为参数名,父节点时为表格名
    parentNodeSelected = false;
    parentNodeSelected = distinguishNodes( filtername );  // 判断是父节点还是子节点发生了变化
    if ( parentNodeSelected )
        fatherNodeClicked( item, filtername );  // 父节点被点击，同步节点状态，显示整张表
    else
        childrenNodeClicked( item, filtername );  // 子节点被点击，显示单列数据，取消点击，删除单列数据
}

// 判断是父节点还是子节点发生了变化
bool MainWindow::distinguishNodes( QString filtername ) {
    for ( int i = 0; i < tablelist.count(); i++ ) {
        if ( QString::compare( filtername, tablelist.at( i ), Qt::CaseSensitive ) )  // 返回0表示相等
            ;
        else {
            return true;
        }
    }
    return false;
}

// 父节点被点击，同步节点状态，显示整张表
void MainWindow::fatherNodeClicked( QTreeWidgetItem* item, QString filtername ) {
    // 遍历 item 结点所有的子结点，设置子父节点同状态
    for ( int i = 0; i < item->childCount(); i++ ) {
        // 找到每个子结点
        QTreeWidgetItem* childItem = item->child( i );
        // 将子结点的选中状态调整为和父结点相同
        childItem->setCheckState( 0, item->checkState( 0 ) );
    }

    if ( item->checkState( 0 ) )  // 被选中，未选中为0，半选1，全选2
    {
        // 父节点不能同时被选中
        for ( int i = 0; i < tablelist.count(); i++ ) {
            if ( QString::compare( filtername, tablelist.at( i ), Qt::CaseSensitive ) )  // 返回0表示相等
            {

                QList< QTreeWidgetItem* > otheritems = ui->treeWidget->findItems( tablelist.at( i ), Qt::MatchFixedString );
                otheritems.at( 0 )->setCheckState( 0, Qt::Unchecked );
            }
            else
                ;
        }
        // 显示整张表格
        qDebug() << "用户选择了父节点" << filtername;
        mymodel->setTable( filtername );
        mymodel->select();
        ui->tableView->setModel( mymodel );
    }
    else  // 未选中
    {
        // 清空tableview
        model->clear();
        ui->tableView->setModel( model );
    }
}


// 子节点被点击，显示单列数据，取消点击，删除单列数据
void MainWindow::childrenNodeClicked( QTreeWidgetItem* item, QString filtername ) {
    int columncount = model->columnCount();
    qDebug() << columncount;                                    // 当前列
    QString filtertable = item->parent()->text( 0 ).trimmed();  // 子节点时为表格名

    if ( item->checkState( 0 ) )  // 选中状态
    {
        qDebug() << "用户选择了子节点" << filtername;
        QString   filtersql = "select %1.`%2` from %1";  // 筛选结果
        QSqlQuery filterquery( dbMYSQL );
        filterquery.exec( filtersql.arg( filtertable ).arg( filtername ) );

        QVector< double >       y;
        QList< QStandardItem* > items;
        while ( filterquery.next() ) {
            QString value = filterquery.value( 0 ).toString();  // 单个数据值
            qDebug() << "单个数据值" << value;
            //                y<<value.toDouble();
            QStandardItem* item = new QStandardItem( value );  // QStandardItem是存储数据的单元格，它存储的是QString
            items << item;
        }
        model->appendColumn( items );                                                    // 插入列
        model->setHorizontalHeaderItem( columncount, new QStandardItem( filtername ) );  // 设置表头

        // draw( filtername, y, columncount );
    }
    else {
        int columncount = model->columnCount();
        qDebug() << "当前有列：" << columncount;
        for ( int j = 0; j < columncount; j++ )  // 遍历每一个单元格(列)
        {
            QVariant currenttext = model->headerData( j, Qt::Horizontal );
            if ( QString::compare( filtername, currenttext.toString(), Qt::CaseSensitive ) )
                ;
            else
                model->takeColumn( j );
        }
        //        QList<QStandardItem *>  removeitems= model->findItems(filtername,Qt::MatchFixedString);//包含项对全树进行搜索
        //        if(removeitems.count())
        //        {
        //            currentcolumn=removeitems.at(0)->column();//当前列
        //            qDebug() <<"删除当前列："<<currentcolumn<<removeitems.at(0)->text();
        //            QModelIndex currentindex=removeitems.at(0)->index();
        //            model->removeColumn(currentcolumn,currentindex);
        //        }
    }
    ui->tableView->setModel( model );  // 表视图，用于显示
}

// 判断表是否存在,存在为真，不存在为假
bool MainWindow::isTableExists( QString& table ) {
    QSqlQuery query( dbMYSQL );
    QString   sql = QString( "show tables;" );  // 查询数据库中是否存在表名
    query.exec( sql );
    while ( query.next() ) {
        QString biaoming = query.value( 0 ).toString().trimmed();
        // qDebug() << "数据库中已存在：" << biaoming;
        if ( QString::compare( biaoming, table, Qt::CaseInsensitive ) )
            ;
        else  // 存在
            return true;
    }
    return false;
}


// 在mysql中创建新表格
void MainWindow::creatNewTable( QAxObject* work_sheet, QString filename ) {
    // 获取工作表一些属性
    // QString    work_sheet_name = work_sheet->property( "Name" ).toString();  // 获取工作表名称
    QString    work_sheet_name = filename;
    QAxObject* used_range      = work_sheet->querySubObject( "UsedRange" );  // 选取当前页面所有已使用单元格
    QAxObject* columns         = used_range->querySubObject( "Columns" );
    int        column_start    = used_range->property( "Column" ).toInt();  // 获取起始列
    int        column_count    = columns->property( "Count" ).toInt();      // 获取列数
    QString    keyType[ column_count ];                                     // 表头数列

    // 获取表头内容
    for ( int i = column_start; i < column_count + column_start; i++ ) {
        QAxObject* cell             = work_sheet->querySubObject( "Cells(int,int)", 1, i );
        QString    value            = cell->dynamicCall( "Value2()" ).toString();
        keyType[ i - column_start ] = value;
        // qDebug() << i - column_start << ":" << keyType[ i - column_start ];  // 打印表头
    }

    // 按表头在MySQL中创建新表
    QString creatsql = QString( "create table %1(" ).arg( work_sheet_name );
    for ( int i = 0; i <= column_count - 1; i++ ) {
        creatsql = creatsql + QString( "%1" ).arg( keyType[ i ] );
        if ( i < column_count - 1 ) {
            creatsql = creatsql + QString( " varchar(255)," );
        }
        else {
            creatsql = creatsql + QString( " varchar(255));" );
        }
    }
    // qDebug() << creatsql;  // 打印创建表MySQL命令
    QSqlQuery creatquery( dbMYSQL );
    if ( creatquery.exec( creatsql ) )
        qDebug() << "成功创建表格：" << work_sheet_name;
    else
        qDebug() << "创建表格" << work_sheet_name << "失败";
    qDebug() << "Error: " << creatquery.lastError().text();
}


// 在表格中插入数据
void MainWindow::InsertData( QAxObject* work_sheet, QString filename ) {
    // QString    work_sheet_name = work_sheet->property( "Name" ).toString();  // 获取工作表名称
    QString    work_sheet_name = filename;
    QAxObject* used_range      = work_sheet->querySubObject( "UsedRange" );  // 选取当前页面所有已使用单元格
    QAxObject* rows            = used_range->querySubObject( "Rows" );
    QAxObject* columns         = used_range->querySubObject( "Columns" );
    int        row_start       = used_range->property( "Row" ).toInt();     // 获取起始行
    int        column_start    = used_range->property( "Column" ).toInt();  // 获取起始列
    int        row_count       = rows->property( "Count" ).toInt();         // 获取行数
    int        column_count    = columns->property( "Count" ).toInt();      // 获取列数
    // qDebug() << "row_count:" << row_count;

    QSqlQuery insertquery( dbMYSQL );
    for ( int i = row_start; i < row_count + row_start; i++ )  // 从行开始
    {
        QString strSql = QString( "insert into %1 values(" ).arg( work_sheet_name );
        for ( int j = column_start; j < column_count + column_start; j++ ) {
            QAxObject* cell  = work_sheet->querySubObject( "Cells(int,int)", i + 1, j );  // 去除第一行表头
            QString    Value = cell->dynamicCall( "Value2()" ).toString();
            strSql           = strSql + QString( "'%1'" ).arg( Value );
            if ( j < column_count ) {
                strSql = strSql + QString( "," );
            }
            else {
                strSql = strSql + QString( ")" );
            }
        }
        // qDebug() << strSql;
        if ( !insertquery.exec( strSql ) )
            qWarning() << "Failed to insert empty rows:" << insertquery.lastError().text();
        // else
        //     qDebug() << "insert successfully.";
    }

    // 验证是否缺少行,不缺少则数据导入成功
    QString resuresql = QString( "select count(1) from %1" ).arg( work_sheet_name );
    insertquery.exec( resuresql );
    insertquery.next();
    if ( insertquery.value( 0 ).toInt() )
        QMessageBox::warning( this, tr( "提示：" ), tr( "数据导入成功" ) );
    else
        QMessageBox::warning( this, tr( "提示：" ), tr( "数据缺失，请重新导入" ) );
}

MainWindow::~MainWindow() {
    delete ui;
}


void MainWindow::on_pushButton_import_clicked() {
    QString filePath = QFileDialog::getOpenFileName( this, tr( "Open Excel file" ), "", tr( "Excel Files (*.xlsx *.xls)" ) );  // 打开选择文件窗口
    if ( filePath.isEmpty() ) {
        qDebug() << "未选择文件";
    }
    else {
        qDebug() << "当前打开的文件路径为" << filePath;  // 显示文件路径
        QFileInfo fileInfo( filePath );
        QString   filename = fileInfo.baseName();  // 获取工作簿（Excel文件）名称，仅含一个Sheet表

        QAxObject* excel = new QAxObject( "Excel.Application" );
        excel->setProperty( "Visible", false );                            // 不显示 Excel 窗体
        QAxObject* work_books = excel->querySubObject( "WorkBooks" );      // 获取工作簿（Excel文件）集合
        work_books->dynamicCall( "Open (const QString&)", filePath );      // 打开已存在的工作簿
        QAxObject* work_book = excel->querySubObject( "ActiveWorkBook" );  // 获取当前工作簿（Excel文件）

        QAxObject* work_sheets = work_book->querySubObject( "Sheets" );     // Sheets也可换用WorkSheets
        int        sheet_count = work_sheets->property( "Count" ).toInt();  // 获取工作表数目
        qDebug() << "当前文件有" << sheet_count << "张sheet";

        for ( int i = 1; i <= sheet_count; i++ )  // 循环操作每张sheet
        {
            QAxObject* work_sheet = work_book->querySubObject( "Sheets(int)", i );
            // QString    work_sheet_name = work_sheet->property( "Name" ).toString();  // 获取工作表名称
            QVariant   visible = work_sheet->dynamicCall( "Visible" );  // 导入可见工作表
            if ( visible.toInt() == -1 )                                // -1表示该工作表是可见的
            {
                qDebug() << "所获取工作簿（Excel文件）名" << filename;

                QString main_sheetName = QString( "main" );
                if ( !isTableExists( main_sheetName ) ) {
                    creatNewTable( work_sheet, QString( "main" ) );  // 若合并表不存在，则创建与打开文件拥有一样表头的Table合并表
                }
                else {
                    // 丢弃'main'表,重新统计数量,创建'main'表
                    QSqlQuery query( dbMYSQL );
                    query.exec( QString( "drop table if exists %1" ).arg( QString( "main" ) ) );
                    qDebug() << "'main' 存在，丢弃，重新统计";
                    creatNewTable( work_sheet, QString( "main" ) );  // 若合并表不存在，则创建与打开文件拥有一样表头的Table合并表
                }

                // 查看数据库中是否有相应表格，没有就创建
                if ( isTableExists( filename ) ) {
                    int ok = QMessageBox::warning( this, tr( "提示：" ),
                                                   tr( "当前数据库中已经存在该表，"
                                                       "确认替换吗？ " ),
                                                   QMessageBox::Yes, QMessageBox::No );
                    if ( ok == QMessageBox::No )
                        ;
                    else  // 替换
                    {
                        QSqlQuery query( dbMYSQL );
                        query.exec( QString( "drop table if exists %1" ).arg( filename ) );
                        creatNewTable( work_sheet, filename );
                        InsertData( work_sheet, filename );
                    }
                }
                else {
                    creatNewTable( work_sheet, filename );
                    InsertData( work_sheet, filename );
                }
            }
        }

        // NOTE 合并所有'立创'开头的表,插入main表中
        QSqlQuery query;  // 查询符合条件的表名，例如选择所有名称以 '立创' 开头的表
        query.prepare( "SELECT table_name FROM information_schema.tables WHERE table_schema = :schema_name AND table_name LIKE :table_pattern" );
        query.bindValue( ":schema_name", "mydatabase" );  // 使用目标数据库名称
        query.bindValue( ":table_pattern", "立创%" );     // 例如选择所有名称以 '立创' 开头的表

        if ( !query.exec() ) {
            qDebug() << "Error fetching table names:" << query.lastError().text();
        }

        // 输出符合条件的所有表名
        while ( query.next() ) {
            QString tableName = query.value( 0 ).toString();
            qDebug() << "Found table:" << tableName;

            // 可以在此处进行后续操作，如选择数据、插入数据等
            // 例如：查询该表的数据
            QSqlQuery selectQuery;
            selectQuery.prepare( QString( "INSERT INTO main SELECT * FROM %1" ).arg( tableName ) );
            if ( selectQuery.exec() ) {
                while ( selectQuery.next() ) {
                    // 假设每个表有 'id', 'name' 字段
                    QString id   = selectQuery.value( "商品编号" ).toString();
                    QString name = selectQuery.value( "名称" ).toString();
                    // qDebug() << "ID:" << id << ", Name:" << name;
                }
                qDebug() << "Data copied successfully!";
            }
            else {
                qDebug() << "Error selecting data from table:" << tableName << selectQuery.lastError().text();
            }
        }

        // NOTE 删除main表中所有列都为空或空字符串的行
        QString tableName = QString( "main" );                                               // 查询全部列名
        QString queryStr  = QString( "SELECT * FROM %1 LIMIT 1" ).arg( QString( "main" ) );  // 查询一行即可
        if ( !query.exec( queryStr ) ) {
            qWarning() << "Failed to execute query:" << query.lastError().text();
            return;
        }
        QSqlRecord  record           = query.record();  // 获取查询结果
        int         columnName_count = record.count();
        QStringList columnNames;
        // qDebug() << "Columns in table" << tableName << ":";
        for ( int i = 0; i < record.count(); ++i ) {
            QString columnName = record.fieldName( i );
            columnNames.insert( columnNames.length(), columnName );
            // qDebug() << columnNames.last();
        }

        QString deleteQuery_str = QString( "DELETE FROM main \n WHERE (%1 IS NULL OR %2 = '')" ).arg( columnNames.at( 0 ) ).arg( columnNames.at( 0 ) );
        for ( int i = 1; i <= columnName_count - 1; i++ ) {
            if ( i < columnName_count - 1 ) {
                deleteQuery_str = deleteQuery_str + QString( "\n AND (%1 IS NULL OR %2 = '')" ).arg( columnNames.at( i ) ).arg( columnNames.at( i ) );
            }
            else {
                deleteQuery_str = deleteQuery_str + QString( " \n AND (%1 IS NULL OR %2 = '');" ).arg( columnNames.at( i ) ).arg( columnNames.at( i ) );
            }
        }
        if ( !query.exec( deleteQuery_str ) )  // 删除空白行
            qWarning() << "Failed to delete empty rows:" << query.lastError().text();
        else
            qDebug() << "Empty rows deleted successfully.";

        // NOTE 丢弃'order_counts'表,重新统计数量
        if ( !query.exec( QString( "drop table if exists %1" ).arg( QString( "order_counts" ) ) ) )  // 如果 'order_counts' 存在，丢弃
            qWarning() << "Failed to drop table order_counts" << query.lastError().text();
        else
            qDebug() << "'order_counts' 存在，丢弃后重新统计";

        QSqlQuery createTableQuery;
        createTableQuery.prepare( "CREATE TABLE IF NOT EXISTS order_counts ("
                                  "商品编号  varchar(255), "
                                  "商品分类 varchar(255), "
                                  "名称 varchar(255), "
                                  "商品型号 varchar(255), "
                                  "封装规格 varchar(255), "
                                  "total_quantity INT, "
                                  "PRIMARY KEY (商品编号));" );  // 创建新表 'order_counts'

        if ( !createTableQuery.exec() ) {
            qDebug() << "Error creating table 'order_counts':" << createTableQuery.lastError().text();
        }
        else {
            qDebug() << "Table 'order_counts' created successfully.";
        }

        QString sumQuery_str = R"(
            INSERT INTO order_counts (商品编号,商品分类,名称,商品型号,封装规格,total_quantity)
            SELECT  商品编号,商品分类,名称,商品型号,封装规格, SUM(购买数量) AS total_quantity
            FROM main
            GROUP BY 商品编号,商品分类,名称,商品型号,封装规格;)";  // 统计相同 [ 商品编号,商品分类,名称,商品型号,封装规格 ] 的数量

        if ( !query.exec( sumQuery_str ) ) {
            qDebug() << "Error inserting data into 'order_counts':" << query.lastError().text();
        }
        else {
            qDebug() << "Data inserted into 'order_counts' successfully.";
        }

        work_book->dynamicCall( "Close()" );  // 关闭工作簿
        excel->dynamicCall( "Quit()" );       // 关闭 excel
        delete excel;
        excel = NULL;
    }
}
