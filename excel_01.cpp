#include "excel_01.h"
#include "ui_excel_01.h"
#include <QTableWidget>
#include <QFileDialog>
#include <QDesktopServices>
#include<QStringList>
#include<QString>
#include<QtDebug>
#include<utility>
#include <QMainWindow>
#include<vector>
#include<QMessageBox>
#include <QAxObject>
#include <QTime>
#include<QVector>
excel_01::excel_01(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::excel_01)
{
    ui->setupUi(this);
    ui->progressBar->hide();




}





void excel_01::on_add_pushButton_clicked()
{
    //添加行
    int crow = ui->tableWidget->currentRow();

    ui->tableWidget->insertRow(crow+1);
}

void excel_01::on_delet_pushButton_clicked()
{
    //删除行
    int crow = ui->tableWidget->currentRow();
    ui->tableWidget->removeRow(crow);
}

void excel_01::on_pushButton_clicked()
{
    //导出
    ui->progressBar->setValue(0);   //设置进度条的值为0
            QString fileName = QFileDialog::getSaveFileName(this,tr("Excle file"),QString("./paper_list.xlsx"),tr("Excel Files(*.xlsx)"));    //设置保存的文件名
            if(fileName != "")
            {    ui->progressBar->show();    //进度条需要在ui文件中加个progressBar控件

                 ui->progressBar->setValue(10);
                 QAxObject *excel = new QAxObject;
                 if(excel->setControl("Excel.Application"))
                 { dynamicCall("SetVisible (bool Visible)",false);
                     excel->setProperty("DisplayAlerts",false);
                     QAxObject *workbooks = excel->querySubObject("WorkBooks");            //获取工作簿集合
                     workbooks->dynamicCall("Add");                                        //新建一个工作簿
                     QAxObject *workbook = excel->querySubObject("ActiveWorkBook");        //获取当前工作簿
                     QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);
                     QAxObject *cell;


                     /*添加Excel表头数据*/
                     for(int i = 1; i <= ui->tableWidget->columnCount(); i++)
                     {
                         cell=worksheet->querySubObject("Cells(int,int)", 1, i);
                         cell->setProperty("RowHeight", 40);
                         cell->dynamicCall("SetValue(const QString&)", ui->tableWidget->horizontalHeaderItem(i-1)->data(0).toString());
                         if(ui->progressBar->value()<=50)
                         {
                             ui->progressBar->setValue(10+i*5);
                         }
                     }


                     /*将form列表中的数据依此保存到Excel文件中*/
                     for(int j = 2; j<=ui->tableWidget->rowCount()+1;j++)
                     {
                         for(int k = 1;k<=ui->tableWidget->columnCount();k++)
                         {
                             cell=worksheet->querySubObject("Cells(int,int)", j, k);
                             if(ui->tableWidget->item(j-2,k-1)!=NULL){

                                 cell->dynamicCall("SetValue(const QString&)",ui->tableWidget->item(j-2,k-1)->text()+ "\t");
                             }
                         }
                         if(ui->progressBar->value()<80)
                         {
                             ui->progressBar->setValue(50+j*5);
                         }
                     }


                     /*将生成的Excel文件保存到指定目录下*/
                     workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(fileName)); //保存至fileName
                     workbook->dynamicCall("Close()");                                                   //关闭工作簿
                     excel->dynamicCall("Quit()");                                                       //关闭excel
                     delete excel;
                     excel=NULL;


                     ui->progressBar->setValue(100);
                     if (QMessageBox::question(NULL,QString::fromUtf8("完成"),QString::fromUtf8("文件已经导出，是否现在打开？"),QMessageBox::Yes|QMessageBox::No)==QMessageBox::Yes)
                     {
                         QDesktopServices::openUrl(QUrl("file:///" + QDir::toNativeSeparators(fileName)));
                     }
                     ui->progressBar->setValue(0);
                     ui->progressBar->hide();
                 }
            }

}

void excel_01::on_pushButton_2_clicked()
{
    //导入
    ui->progressBar->setValue(0);   //设置进度条的值为0
    QString path = QFileDialog::getOpenFileName(this,"open",
                                              "../","execl(*.xlsx)");
    //指定父对象（this），“open”具体操作，打开，“../”默认，之后可以添加要打开文件的格式
    if(path.isEmpty()==false)
    {
        //文件对象
        QFile file(path);
        //打开文件,默认为utf8变量，
        bool flag = file.open(QIODevice::ReadOnly);
        if(flag == true)//打开成功
        {
            ui->progressBar->show();    //进度条需要在ui文件中加个progressBar控件

            ui->progressBar->setValue(10);
            QAxObject *excel = new QAxObject(this);//建立excel操作对象
            excel->setControl("Excel.Application");//连接Excel控件
            excel->setProperty("Visible", false);//不显示窗体看效果
            excel->setProperty("DisplayAlerts", false);//不显示警告看效果
            /*********获取COM文件的一种方式************/
            QAxObject *workbooks = excel->querySubObject("WorkBooks");
            //获取工作簿(excel文件)集合
            workbooks->dynamicCall("Open(const QString&)", path);//path至关重要，获取excel文件的路径
            //打开一个excel文件
            QAxObject *workbook = excel->querySubObject("ActiveWorkBook");
            QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)",1);//访问excel中的工作表中第一个单元格
            QAxObject *usedRange = worksheet->querySubObject("UsedRange");//sheet的范围
            /*********获取COM文件的一种方式************/
            //获取打开excel的起始行数和列数和总共的行数和列数
            int intRowStart = usedRange->property("Row").toInt();//起始行数
            int intColStart = usedRange->property("Column").toInt(); //起始列数
            QAxObject *rows, *columns;
            rows = usedRange->querySubObject("Rows");//行
            columns = usedRange->querySubObject("Columns");//列
            int intRow = rows->property("Count").toInt();//行数
            int intCol = columns->property("Count").toInt();//列数
            //起始行列号
            //qDebug()<<intRowStart;
            //qDebug()<<intColStart;
            //行数和列数
            //qDebug()<<intRow;
            //qDebug()<<intCol;
            int a,b;
            a=intRow-intRowStart+1,b=intCol-intColStart+1;
            QByteArray text[a][b];
            QString exceldata[a][b];
            int coerow=0,coecol=0;

            for (int i = intRowStart; i < intRowStart + intRow; i++,coerow++)
                {
                    coecol=0;//务必是要恢复初值的
                    for (int j = intColStart; j < intColStart + intCol; j++,coecol++)
                    {
                        auto cell = excel->querySubObject("Cells(Int, Int)", i, j );
                        QVariant cellValue = cell->dynamicCall("value");
                        text[coerow][coecol]=cellValue.toByteArray();//QVariant转换为QByteArray
                        exceldata[coerow][coecol]=QString(text[coerow][coecol]);//QByteArray转换为QString
                        if(ui->progressBar->value()<=60)
                        {
                            ui->progressBar->setValue(10+i*5);
                        }
                        //qDebug()<<exceldata[coerow][coecol]<<coerow<<" "<<coecol;
                    }
                }
            ui->tableWidget->setRowCount(a-1);
            for(int i=1;i<a;++i)
                for(int j=0;j<b;++j)
                {
                    ui->tableWidget->setItem(i-1, j, new QTableWidgetItem(exceldata[i][j]));
                    if(ui->progressBar->value()<=80)
                    {
                        ui->progressBar->setValue(60+i*5);
                    }
                }


            workbook->dynamicCall( "Close(Boolean)", false );
            excel->dynamicCall( "Quit(void)" );
            delete excel;
            ui->progressBar->setValue(100);
            QMessageBox::warning(this,tr("读取情况"),tr("读取完成！"),QMessageBox::Yes);

            ui->progressBar->hide();
            ui->progressBar->setValue(0);
        }
        file.close();
    }
}


bool cum(std::pair<int,double>x,std::pair<int,double>y)
{
    return x.second > y.second;
}
void excel_01::sortByColumn(int n)
    {
        if(n!=14)
            return;
        if(SortUporDown)
        {
            ui->tableWidget->sortItems(n, Qt::AscendingOrder);
            SortUporDown=false;
        }
        else
        {
            ui->tableWidget->sortItems(n, Qt::DescendingOrder);
            SortUporDown=true;
        }
    }
