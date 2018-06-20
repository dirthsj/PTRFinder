#include "searchresults.h"
#include "ui_searchresults.h"

#include <iostream>

SearchResults::SearchResults(int ptrNumber, QAxObject* workbooks, QWidget *parent) :
    workbooks(workbooks),
    ptrNumber(ptrNumber),
    QWidget(parent),
    ui(new Ui::SearchResults)
{
    ui->setupUi(this);
    int th = ptrNumber / 1000 * 1000;
    int hu = ptrNumber / 100 * 100;
    QString thousands = QString::number( th ) + "-" + QString::number( th + 999 );
    QString hundreds = QString::number( hu ) + "-" + QString::number( hu + 99 );
    QString actual = QString::number( ptrNumber );
    QDir inProcess, completed, cancelled;
    inProcess.setPath( "P:\\Testing\\PTRs\\InProcess\\" + thousands + "\\" + hundreds + "\\" + actual );
    completed.setPath( "P:\\Testing\\PTRs\\Completed\\" + thousands + "\\" + hundreds + "\\" + actual );
    cancelled.setPath( "P:\\Testing\\PTRs\\Cancelled\\" + thousands + "\\" + hundreds + "\\" + actual );
    if( inProcess.exists() ){
        getExcelFilePath(inProcess);
        ui->ptrLabel->setText( "Found PTR " + actual + ".\nStatus: In Process\nWhat would you like to do?" );
        path = inProcess;
    }else if( completed.exists() ){
        getExcelFilePath(completed);
        ui->ptrLabel->setText( "Found PTR " + actual + ".\nStatus: Completed\nWhat would you like to do?" );
        path = completed;
    }else if( cancelled.exists() ){
        getExcelFilePath(cancelled);
        ui->ptrLabel->setText( "Found PTR " + actual + ".\nStatus: Cancelled\nWhat would you like to do?" );
    }else{
        ui->ptrLabel->setText( "Could not find PTR " + actual );
        ui->excelButton->setDisabled(true);
        ui->folderButton->setDisabled(true);
    }
    ui->excelButton->setEnabled( excelFilePath.exists() );
    this->setWindowTitle(actual);
}

SearchResults::~SearchResults()
{
    delete ui;
}

void SearchResults::on_closeButton_clicked()
{
    this->close();
}

void SearchResults::on_folderButton_clicked()
{
    if( QDesktopServices::openUrl( QUrl::fromLocalFile( path.absolutePath() ) ) ){
        this->close();
    }
}

void SearchResults::on_excelButton_clicked()
{
    if(QDesktopServices::openUrl( QUrl::fromLocalFile( excelFilePath.absoluteFilePath() ))){
        this->close();
    }
}

void SearchResults::getExcelFilePath( QDir searchPath ){
    for( QFileInfo file : searchPath.entryInfoList() ){
        if( file.suffix() == "xls" ){
           QAxObject* workbook = workbooks->querySubObject( "Open(const QString&)", file.absoluteFilePath() );
           QAxObject* sheets = workbook->querySubObject( "Worksheets" );
           int count = sheets->property("Count").toInt();
           for( int i = 1; i <= count; i++){
               QAxObject* sheet = sheets->querySubObject("Item( int )", i );
               QAxObject* ptrNumberCell = sheet->querySubObject( "Cells( int, int )", 2, 18 ); // R2
               QVariant ptrNumberCellValue = ptrNumberCell->dynamicCall( "Value()" );
               if( ptrNumberCellValue == ptrNumber ){
                   excelFilePath = file;
                   break;
               }
           }
           workbook->dynamicCall("Close()");
        }
    }
}
