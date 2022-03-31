#ifndef EXCEL_01_H
#define EXCEL_01_H

#include <QDialog>

namespace Ui {
class excel_01;
}

class excel_01 : public QDialog
{
    Q_OBJECT

public:
    explicit excel_01(QWidget *parent = nullptr);
    //virtual ~excel_01();
    //bool cum(std::pair<int,double>x,std::pair<int,double>y);

protected:
    Ui::excel_01 *ui;
    std::vector<std::pair<int,double>> src;
    bool SortUporDown=true;






protected slots:

    virtual void on_count_pushButton_clicked()=0;
private slots:
    void sortByColumn(int n);
    void on_add_pushButton_clicked();
    void on_delet_pushButton_clicked();
    void on_pushButton_clicked();
    void on_pushButton_2_clicked();

private:

};


bool cum(std::pair<int,double>x,std::pair<int,double>y);


#endif // EXCEL_01_H
