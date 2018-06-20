#include "mainwindow.h"
#include "ui_mainwindow.h"

#include "searchresults.h"

#include <iostream>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    excel = new QAxObject( "Excel.Application", 0 );
    workbooks = excel->querySubObject( "Workbooks" );
}

MainWindow::~MainWindow()
{
    excel->dynamicCall("Quit()");
    delete ui;
}

void MainWindow::on_searchButton_clicked()
{
    QString ptrNumberString = ui->ptrNumber->text();
    int ptrNumber = ptrNumberString.toInt();
    std::cout << ptrNumber << std::endl;
    SearchResults* results = new SearchResults( ptrNumber, workbooks );
    results->show();
}

void MainWindow::on_ptrNumber_returnPressed()
{
    ui->searchButton->click();
}
