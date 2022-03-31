#include "excel_01.h"
#include "ui_excel_01.h"
#include"exc02.h"
#include<QDebug>
#include<QMessageBox>


exc02::exc02(QWidget *parent) :
    excel_01(parent)
{
    setWindowTitle("个人，集体自编动作评分系统");
    ui->tableWidget->setRowCount(5);
    ui->tableWidget->setColumnCount(15);
    ui->tableWidget->setAlternatingRowColors(true);

    ui->tableWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);

    QStringList headers;
    headers << QStringLiteral("签号") << QStringLiteral("姓名") <<QStringLiteral("代表队")<< QStringLiteral("体式A")<<QStringLiteral("体式B")
                  << QStringLiteral("体式C")<<QStringLiteral("展示A")<< QStringLiteral("展示B")<<QStringLiteral("展示C")
                     << QStringLiteral("难度A")<<QStringLiteral("难度B")<< QStringLiteral("难度C")<<QStringLiteral("裁判长扣分")
                        << QStringLiteral("成绩")<<QStringLiteral("排名");
    ui->tableWidget->setHorizontalHeaderLabels(headers);
    connect(ui->tableWidget->horizontalHeader(), SIGNAL(sectionClicked(int)),
    this, SLOT(sortByColumn(int)));
}

void exc02::on_count_pushButton_clicked()
{
    src.clear();
    int count=0;
    for(int i=0;i<ui->tableWidget->rowCount();i++)
    {

        std::map<double,int>item;
        double sorce=0;
        QString str="";
        for(int j=3;j<ui->tableWidget->columnCount()-2;j++)

        {

            try
            {
                bool flag=false;

                if(ui->tableWidget->item(i,j)!=nullptr)
                {
                    if(ui->tableWidget->item(i,j)->text()!=""){

                        str = ui->tableWidget->item(i,j)->text();
                        double d=str.toDouble(&flag);
                        if(!flag)
                            throw -1;
                        if(j==12)
                        {

                            sorce-=d;
                        }
                        else if(j>=9&&j<12)
                        {
                            item[d]++;

                        }
                        else
                            sorce += d/3.0;

                    }

                }


            }
            catch(int)
            {

                ui->tableWidget->setItem(i, j, new QTableWidgetItem(QString("")));
            }
        }


        for(auto i:item)
        {

            if(i.second>=2)
            {
                sorce+=i.first;
                break;
            }


        }
        str =QString::number(sorce, 'f', 2);
        ui->tableWidget->setItem(i, 13, new QTableWidgetItem(str));


        auto p=std::pair<int,double>(i,sorce);
        src.push_back(p);
    }

    std::sort(src.begin(),src.end(),cum);
    count=1;
    for(auto i:src)
    {
        //qDebug()<<"aaaaa";
        QTableWidgetItem *it = new QTableWidgetItem;
        it->setData(Qt::DisplayRole, count++);
        //DisplayRole!!!!
        ui->tableWidget->setItem(i.first,14,it);

    }

}
