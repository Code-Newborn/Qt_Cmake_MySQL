#include "mainwindow.h"
#include "./ui_mainwindow.h"

#include <QtSql/QSqlDatabase>
#include <QtSql/QSqlDriver>
#include <QtDebug>
#include <QSqlError>
#include <QSqlQuery>
#include <QFile>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    // 输出可用数据库
    qDebug() << "available drivers:";
    QStringList drivers = QSqlDatabase::drivers();
    foreach( QString driver, drivers )
        qDebug() << driver;

    // 删除已存在的表
    QString filePath = "C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/t.csv";
    QFile file(filePath);
    if (file.exists()) {
        file.remove();
    }


    QSqlDatabase db = QSqlDatabase::addDatabase( "QMYSQL" );  // 添加驱动
    db.setPort( 3306 );
    db.setHostName( "127.0.0.1" );       // ip地址
    db.setDatabaseName( "mydatabase" );  // 数据库名 在MySQL 8.0 Command Line Client中使用命令SHOW DATABASES;查询
    db.setUserName( "root" );            // 用户名
    db.setPassword( "Cz1253709179." );   // 密码
    if ( db.open() ) {
        qDebug() << "open successful";  // 如果连接成功打印 open successful
    }
    else {
        qDebug() << "error" << db.lastError().text();  // 连接失败打印error信息
    }

    QSqlQuery query;

    // 删表语句
    bool success = query.exec( "DROP TABLE IF EXISTS city" );
    if ( !success ) {
        qDebug() << "Error deleting table:" << query.lastError();
    }
    else {
        qDebug() << "Table deleted successfully";
    }


    // 建表语句
    QString   createTableSql = "CREATE TABLE IF NOT EXISTS city ("
                               "Name VARCHAR(255) NOT NULL PRIMARY KEY, "
                               "CountryCode VARCHAR(255) NOT NULL, "
                               "District VARCHAR(255) NOT NULL UNIQUE, "
                               "Population INT NOT NULL)";
    if ( !query.exec( createTableSql ) ) {
        qDebug() << "Error: 创建表失败," << query.lastError().text();
    }
    else {
        qDebug() << "创建表成功";
    }

    db.exec( "SET NAMES 'UTF8'" );  // 防止插入的中文数据为乱码

    // 插入语句
    QString sqlStr = "insert into city(Name,CountryCode,District,Population)values(:Name,:CountryCode,:District,:Population);";
    query.prepare( sqlStr );
    query.bindValue( ":Name", "广东" );
    query.bindValue( ":CountryCode", "GUA" );
    query.bindValue( ":District", "广州" );
    query.bindValue( ":Population", 123301 );
    if ( query.exec() ) {
        qDebug() << "insert success!";
    }
    else {
        qDebug() << "insert failed:" << query.lastError().text();
    }

    // 保证C:\ProgramData\MySQL\MySQL Server XX.XX\my.ini文件里的secure-file-priv字段所设置文件夹路径为允许导出路径
    if ( query.exec( "SELECT * FROM mydatabase.city INTO OUTFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/t.csv';" ) ) {
        qDebug() << "导出成功：";
    }
    else {
        qDebug() << "导出失败：" << query.lastError();
    }


    // 执行查询
    if ( query.exec( "SELECT * FROM mydatabase.city" ) ) {  // 替换为你的SQL查询
        // 遍历查询结果
        while ( query.next() ) {
            QString field1 = query.value( 0 ).toString();  // 假设第一个字段是字符串类型
            int     field2 = query.value( 3 ).toInt();     // 假设第二个字段是整型
            qDebug() << field1 << field2;
        }
    }
    else {
        qDebug() << "查询失败：" << query.lastError();
    }

    // 关闭数据库连接
    db.close();
}

MainWindow::~MainWindow()
{
    delete ui;
}

