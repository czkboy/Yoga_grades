#ifndef EXC01_H
#define EXC01_H
#include "excel_01.h"
#include "ui_excel_01.h"
#include <QDialog>

namespace Ui {
class exc01;
}

class exc01 : public excel_01
{
    Q_OBJECT

public:
    explicit exc01(QWidget *parent = nullptr);
    //virtual ~exc01();

private slots:
    void on_count_pushButton_clicked();

private:
//    Ui::exc01 *ui;
};
#endif // EXC01_H
