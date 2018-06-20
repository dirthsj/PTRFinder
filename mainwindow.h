#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QAxObject>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_searchButton_clicked();

    void on_ptrNumber_returnPressed();

private:
    Ui::MainWindow *ui;
    QAxObject* excel;
    QAxObject* workbooks;
};

#endif // MAINWINDOW_H
