#include "mainwindow.h"
#include "ui_mainwindow.h"
#include"exc01.h"
#include"exc02.h"
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_action_triggered()
{
    excel_01 *_excel = new exc02;
    _excel->show();

}

void MainWindow::on_action_2_triggered()
{
    excel_01 *_excel = new exc01;
    _excel->show();

}
