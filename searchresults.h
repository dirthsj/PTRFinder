#ifndef SEARCHRESULTS_H
#define SEARCHRESULTS_H

#include <QWidget>
#include <QDir>
#include <QDesktopServices>
#include <QUrl>
#include <QAxObject>
#include <QTextStream>

namespace Ui {
class SearchResults;
}

class SearchResults : public QWidget
{
    Q_OBJECT

public:
    explicit SearchResults(int ptrNumber, QAxObject* workbooks, QWidget *parent = 0);
    ~SearchResults();

private slots:
    void on_closeButton_clicked();

    void on_folderButton_clicked();

    void on_excelButton_clicked();

private:
    Ui::SearchResults *ui;
    int ptrNumber;
    QAxObject* workbooks;
    QDir path;
    QString originator;
    QFileInfo excelFilePath;

    void getExcelFilePath( QDir searchPath );
};

#endif // SEARCHRESULTS_H
