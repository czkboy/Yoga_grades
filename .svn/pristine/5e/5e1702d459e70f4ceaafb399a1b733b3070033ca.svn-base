#include "excel_01.h"
#include "ui_excel_01.h"
#include"exc01.h"
#include<QDebug>
#include<QMessageBox>


exc01::exc01(QWidget *parent) :
    excel_01(parent)
{
    setWindowTitle("集体规定动作评分系统");
    ui->tableWidget->setRowCount(5);
    ui->tableWidget->setColumnCount(15);
    ui->tableWidget->setAlternatingRowColors(true);
    //ui->tableWidget->setSelectionBehavior(QAbstractItemView::SelectRows);
    ui->tableWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);

    QStringList headers;
    headers << QStringLiteral("签号") << QStringLiteral("姓名") <<QStringLiteral("代表队") <<QStringLiteral("裁判1")<<QStringLiteral("裁判2")
               << QStringLiteral("裁判3")<<QStringLiteral("裁判4")<< QStringLiteral("裁判5")<<QStringLiteral("裁判6")
                  << QStringLiteral("裁判7")<<QStringLiteral("裁判8")<< QStringLiteral("裁判9")<<QStringLiteral("裁判长扣分")
                     << QStringLiteral("成绩")<<QStringLiteral("排名");
    ui->tableWidget->setHorizontalHeaderLabels(headers);

    connect(ui->tableWidget->horizontalHeader(), SIGNAL(sectionClicked(int)),
    this, SLOT(sortByColumn(int)));

}

void exc01::on_count_pushButton_clicked()
{

    src.clear();
    int count=0;
    for(int i=0;i<ui->tableWidget->rowCount();i++)
    {
        std::vector<double>ve;
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
                        else
                        {
                            ve.push_back(d);
                        }

                    }

                }


            }
            catch(int)
            {

                ui->tableWidget->setItem(i, j, new QTableWidgetItem(QString("")));
            }
        }
        std::sort(ve.begin(),ve.end());
        if(ve.size()>4)
        {
            //qDebug()<<"acc="<<std::accumulate(ve.begin()+2,ve.end()-2,0);
            sorce+=std::accumulate(ve.begin()+2,ve.end()-2,0.0)/double(ve.size()-4);
        }
        //qDebug()<<"sorce=="<<sorce;
        str =QString::number(sorce, 'f', 2);
        ui->tableWidget->setItem(i, 13, new QTableWidgetItem(str));


        auto p=std::pair<int,double>(i,sorce);
        src.push_back(p);
    }

    std::sort(src.begin(),src.end(),cum);
    count=1;

    for(auto i:src)
    {
        QTableWidgetItem *it = new QTableWidgetItem;
        it->setData(Qt::DisplayRole, count++);
        //DisplayRole!!!!
        ui->tableWidget->setItem(i.first,14,it);
    }

}
