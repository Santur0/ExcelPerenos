#include "mainwindow.h"
#include "ui_mainwindow.h"
//  получаем и проверяем пути!
QStringList MainWindow::getFilesFolder()
{
        QStringList list;
        list.empty();
        QDir dir(ui->oldPath->text());
        if(dir.exists())
        {
             list= dir.entryList(QDir::Files | QDir::Dirs | QDir::NoDotAndDotDot);
            for(auto &el:list)
             {
                 el.push_front(ui->oldPath->text()+"/");
             }
            fileFolderStatus->setText("Папка Открыт -"+QString::number(list.size()));
             buttonsEnabled(true);
        }
        else{
            buttonsEnabled(false);
            fileFolderStatus->setText("!!!! Папка НЕ Открыт !!!!! "+QString::number(list.size()));
        }

        ui->statusbar->addPermanentWidget(fileFolderStatus);

        QStringList result;

        for(const auto &el:list)
        {
            QFileInfo info(el);
            if(info.suffix()=="xlsx"|| info.suffix()=="xls"|| info.suffix()=="xlsb"|| info.suffix()=="xltx"|| info.suffix()=="xlt")
            {
                result.push_back(el);
            }
        }
        ui->filebox->setMaximum(result.size());
        ui->filebox->setMinimum(1);
       return result;
}

//Конструктор Класса
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    ui->buttonSave->setEnabled(false);

    month<<"январь"<<"февраль"<<"март"<<"апрель"
    <<"май"<<"июнь"<<"июль"<<"август"
    <<"сентябрь"<<"октябрь"<<"ноябрь"<<"декабрь";

    settings=new QSettings(this);


    QRect temp=settings->value("geometry",QRect(200,200,800,550)).toRect();
    this->setGeometry(temp);

    this->setWindowTitle("Excel");
    this->setWindowIcon(QIcon(":/icon/icon/iconExcel.png"));

    day=settings->value("day",1).toInt();
    currentSheet=settings->value("currentSheet",2).toInt();
    currentFileIndex=settings->value("currentFileIndex",0).toInt();

    ui->index1_1->setText(settings->value(ui->index1_1->objectName()).toString());
    ui->index2_1->setText(settings->value(ui->index2_1->objectName()).toString());
    ui->index3_1->setText(settings->value(ui->index3_1->objectName()).toString());
    ui->index4_1->setText(settings->value(ui->index4_1->objectName()).toString());
    ui->index5_1->setText(settings->value(ui->index5_1->objectName()).toString());
    ui->index5_1->setText(settings->value(ui->index6_1->objectName()).toString());

    ui->index1_2->setText(settings->value(ui->index1_2->objectName()).toString());
    ui->index2_2->setText(settings->value(ui->index2_2->objectName()).toString());
    ui->index3_2->setText(settings->value(ui->index3_2->objectName()).toString());
    ui->index4_2->setText(settings->value(ui->index4_2->objectName()).toString());
    ui->index5_2->setText(settings->value(ui->index5_2->objectName()).toString());
    ui->index5_2->setText(settings->value(ui->index6_2->objectName()).toString());


    ui->indexLineTime->setText(settings->value(ui->indexLineTime->objectName()).toString());
    ui->lineIndexNums->setText(settings->value(ui->lineIndexNums->objectName()).toString());

    ui->newSheetBox->setValue(settings->value("newSheetBox").toInt());

    ui->dayBox->setValue(day);
    ui->filebox->setValue(currentFileIndex);

    ui->oldPath->setText(settings->value("oldPath","C:\\Новая Папка").toString());

    ui->pathNewExcelFile->setText(settings->value("newPath","C:\\Новая Папка\\Новая Папка").toString());

    this->setStatusBar(ui->statusbar);

    fileFolderStatus=new QLabel();
    excelFileStatus=new QLabel();

    fileFolderStatus->setStyleSheet("color:(255,255,255);");
    excelFileStatus->setStyleSheet("color:(255,255,255);");

    QStringList listFolder=getFilesFolder();

    if(!listFolder.isEmpty())
    {
        excel=new Excel(listFolder[currentFileIndex],this->currentSheet);
        buttonsEnabled(true);

        //qDebug()<<"Папка НЕ Пуста";
    }
    else{

        //qDebug()<<"Папка Пуста";
       buttonsEnabled(false);
    }

    QFile file(ui->pathNewExcelFile->text());
    label_currentFileOpen=new QLabel();

    if(file.exists()){
        newExcel=new Excel(ui->pathNewExcelFile->text(),ui->newSheetBox->value());
        //newExcel->setVisible(ui->checkBox->isChecked());
        row=newExcel->getExcelSize().row;
        isNewFileOpen=true;
        if(!listFolder.isEmpty()){
            buttonsEnabled(true);
        }
        label_currentFileOpen->setText("Файл Открыт");

        ui->newSheetBox->setMaximum(newExcel->getCountSheet());
        ui->newSheetBox->setMinimum(1);
    }
    else{
        isNewFileOpen=false;
        buttonsEnabled(false);
        label_currentFileOpen->setText("Файл Не Открыт!!");
    }
    ui->statusbar->addPermanentWidget(label_currentFileOpen);

    currentTimeBoxInit();
}
//Деструктор
MainWindow::~MainWindow()
{
    settings->setValue("geometry",this->geometry());
    settings->setValue("day",day);
    settings->setValue("currentSheet",currentSheet);
    settings->setValue("currentFileIndex",currentFileIndex);

    settings->setValue(ui->index1_1->objectName(),ui->index1_1->text());
    settings->setValue(ui->index2_1->objectName(),ui->index2_1->text());
    settings->setValue(ui->index3_1->objectName(),ui->index3_1->text());
    settings->setValue(ui->index4_1->objectName(),ui->index4_1->text());
    settings->setValue(ui->index5_1->objectName(),ui->index5_1->text());
    settings->setValue(ui->index6_1->objectName(),ui->index6_1->text());

    settings->setValue(ui->index1_2->objectName(),ui->index1_2->text());
    settings->setValue(ui->index2_2->objectName(),ui->index2_2->text());
    settings->setValue(ui->index3_2->objectName(),ui->index3_2->text());
    settings->setValue(ui->index4_2->objectName(),ui->index4_2->text());
    settings->setValue(ui->index5_2->objectName(),ui->index5_2->text());
    settings->setValue(ui->index6_2->objectName(),ui->index6_2->text());

    settings->setValue(ui->lineIndexNums->objectName(),ui->lineIndexNums->text());
    settings->setValue(ui->indexLineTime->objectName(),ui->indexLineTime->text());

    if(isNewFileOpen)
    {
        newExcel->saveExcel();
        newExcel->close();
        delete newExcel;
    }
    excel->close();
    delete excel;
    delete ui;
}


// При Нажатии на кнопку "Выбрать Старый Путь"
void MainWindow::on_buttonChooseOldPath_clicked()
{
    QString oldPath=QFileDialog::getExistingDirectory();
    ui->oldPath->setText(oldPath);
    settings->setValue("oldPath",ui->oldPath->text());
    QStringList listFolder=getFilesFolder();

    if(!listFolder.isEmpty())
    {
        excel=new Excel(listFolder[currentFileIndex],this->currentSheet);
        buttonsEnabled(true);
    }
    else{
        buttonsEnabled(false);
    }

}

// При нажатии на кнопку "Выбрать Новый Путь"
void MainWindow::on_buttonNewExcelFile_clicked()
{
    QString newPath=QFileDialog::getOpenFileName(this,"Выберите Файл Excel!","C:\\","Excel File(*.xlsx)");
    ui->pathNewExcelFile->setText(newPath);
    settings->setValue("newPath",ui->pathNewExcelFile->text());

    QFile file(newPath);
    label_currentFileOpen=new QLabel();

    if(file.exists()){
        newExcel=new Excel(ui->pathNewExcelFile->text(),ui->newSheetBox->value());
        row=newExcel->getExcelSize().row;
        isNewFileOpen=true;
        label_currentFileOpen->setText("Файл Открыт");

        ui->newSheetBox->setMaximum(newExcel->getCountSheet());
        ui->newSheetBox->setMinimum(1);

    }
    else{
        isNewFileOpen=false;
        label_currentFileOpen->setText("Файл Не Открыт!!");
    }
}

// При нажатии на кнопку "Проверить"
void MainWindow::on_buttonChek_clicked()
{

    currentTimeBoxInit();
    getLists();

    QVector<QLabel*> labelLine;
    labelLine<<ui->labelLine1<<ui->labelLine2<<ui->labelLine3<<ui->labelLine4<<ui->labelLine5<<ui->labelLine6;

    QStringList listFolder=getFilesFolder();

    label_currentFileOpen->setText(listFolder[currentFileIndex]);

    excel->setSheet(currentSheet);

    ui->currentTime->setText(excel->getNameSheet());

    QVector<QVariant> days=excel->getColVector(ui->lineIndexNums->text());

    for(int i=0;i<days.size();++i)
    {
        if ( days[i].toInt()==day )
        {
            ui->currentDay->setText(QString::number(day));
            for(int j=0;j<lineEdist.size();++j)
            {

                QString temp=lineEdist[j]->text()+QString::number(i+1);
                labelLine[j]->setText(temp);
                lineTest[j]->setText(excel->WriteData(temp).toString());
            }
        }
    }

    QStringList words=excel->WriteData("A1").toString().toLower().split(" ");

    int monthNum=0;
    int currentYear=0;

    for(int i=0;i<month.size();++i)
    {
        for(int j=0;j<words.size();++j)
        {
            if(month[i]==words[j]){
                monthNum=i+1;
            }

            if(words[j].toInt()>1500&&words[j].toInt()<2999)
            {
                currentYear=words[j].toInt();
            }
        }
    }

    QString strMonth=QString::number(monthNum);

    if(monthNum<10){
        strMonth.push_front("0");
    }

    if(ui->currentDay->text().toInt()<10){
        ui->currentDay->setText("0"+ui->currentDay->text());
    }

    QString strYear=QString::number(currentYear);

    QString currentDateTime=ui->currentDay->text()+"."+strMonth+"."+strYear;
    ui->currentDay->setText(currentDateTime);

    ui->buttonSave->setEnabled(true);

}
// При нажатии на кнопку "Сохранить И перейти на Следующий период!"
void MainWindow::on_buttonSave_clicked()
{
    if(isNewFileOpen)
    {
        newExcel->setSheet(ui->newSheetBox->value());
        ExcelSize sizeExcelNew=newExcel->getExcelSize();

        //qDebug()<<sizeExcelNew.row;

        QString temp="";

        if(row!=1){
            QDate test=newExcel->WriteData(ui->lineIndexNums->text()+QString::number(row-1)).toDate();

            //qDebug()<<test.toString();

            if(!test.isValid())
            {
                //qDebug()<<"row++";
                row++;
            }
        }

        bool isYes=false;

        for(int i=0;i<lineTest.size();++i)
        {
            if(!lineTest[i]->text().isEmpty()){
                isYes=true;
            }
        }

        QVector<QVariant> vc=newExcel->getRowVector(row);
        bool tmp=false;
        //qDebug()<<row;
        for(int i=0;i<vc.size();++i)
        {
            //qDebug()<<"Val " <<vc[i];
            if(vc[i]!=""){
                tmp=true;
            }
        }

        if(tmp==true){
            row++;
        }

        if(isYes){

            QVector<QVariant> test;
            for(int i=0;i<lineLastEdits.size();++i)
            {
                if(!lineLastEdits[i]->text().isEmpty()){
                    temp=lineLastEdits[i]->text()+QString::number(row);
                    newExcel->setData(temp,"'"+lineTest[i]->text());
                }
            }

            if(!ui->currentDay->text().isEmpty())newExcel->setData("A"+QString::number(row),ui->currentDay->text());

            QString currentTime="";

            switch(currentSheet)
            {
            case 2:
                currentTime="0:00";
                ui->currentTimeBox->setCurrentIndex(0);
                break;
            case 3:
                currentTime="6:00";
                ui->currentTimeBox->setCurrentIndex(1);
                break;
            case 4:
                currentTime="12:00";
                ui->currentTimeBox->setCurrentIndex(2);
                break;
            case 5:
                currentTime="18:00";
                ui->currentTimeBox->setCurrentIndex(3);
                break;
            }

            newExcel->setData(ui->indexLineTime->text()+QString::number(row),currentTime);
        }else{
            row--;
        }
    newExcel->saveExcel();
    row++;

    QStringList listFolder=getFilesFolder();

    if(kolSheet==currentSheet){
        currentSheet=1;
        excel->setSheet(currentSheet);
        qDebug()<<"Day++";
        day++;
    }

    if (day==32)
    {
        qDebug()<<"Day = 32";
        currentFileIndex++;

        //Если КОНЕЦ
        if(currentFileIndex==listFolder.size()){
            QMessageBox message;
            message.setText("Все Файлы Записаны!");
            message.setIcon(QMessageBox::Information);
            message.exec();
            timer->stop();
            progress->setValue(0);
            currentFileIndex=0;
            day=1;
            currentSheet=1;
            ui->dayBox->setValue(day);
            ui->filebox->setValue(currentFileIndex);
            buttonsEnabled(false);
        }
        else{
             excel->setPathFile(listFolder[currentFileIndex]);
        }
        ///////
        ui->currentDay->setText("");
        day=1;
        currentSheet=1;
        }

    currentSheet++;
        ui->filebox->setValue(currentFileIndex+1);
    }

    ui->dayBox->setValue(day);
}

// +/-  доступность конопок
void MainWindow::buttonsEnabled(bool is)
{
    ui->buttonChek->setEnabled(is);
    ui->buttonSave->setEnabled(is);
}

void MainWindow::getLists()
{
    lineEdist.clear();
    lineTest.clear();
    lineLastEdits.clear();

    lineEdist<<ui->index1_1<<ui->index2_1<<ui->index3_1<<ui->index4_1<<ui->index5_1<<ui->index6_1;
    lineTest<<ui->line1<<ui->line2<<ui->line3<<ui->line4<<ui->line5<<ui->line6;
    lineLastEdits<<ui->index1_2<<ui->index2_2<<ui->index3_2<<ui->index4_2<<ui->index5_2<<ui->index6_2;
}

// При клике на кнопку "Сохранить Изменения"
void MainWindow::on_pushButton_clicked()
{
    if(isNewFileOpen)
    {
        QStringList list=getFilesFolder();
        if(!list.isEmpty()){
            QMessageBox message;
            message.setText("Вы Точно Хотите Сохранить Изменения?");
            message.setIcon(QMessageBox::Information);
            message.addButton(QMessageBox::Yes);
            message.addButton(QMessageBox::No);
            bool isSave=message.exec();
            if(isSave)
            {
                day=ui->dayBox->value();
                currentFileIndex=ui->filebox->value()-1;
                excel->setPathFile(list[currentFileIndex]);
                setCurrentTimeBox(ui->currentTimeBox->currentIndex());
                settings->setValue("newSheetBox",ui->newSheetBox->value());
                newExcel->setSheet(ui->newSheetBox->value());
             }
        }
    }
}

void MainWindow::on_pushButton_2_clicked()
{
    if(ui->buttonChek->isEnabled()&&ui->buttonSave->isEnabled()){
    main.setWindowTitle("Авто Выполнение!");

    main.setWindowIcon(QIcon(":/icon/icon/iconExcel.png"));

    main.setGeometry(this->x(),this->y(),400,100);
    main.setMaximumHeight(200);
    main.setMaximumWidth(400);

    QVBoxLayout *layoutV=new QVBoxLayout();
    QHBoxLayout *layoutH=new QHBoxLayout();

    QLabel *label=new QLabel("Количество Повторений!",&main);
    label->setAlignment(Qt::AlignBottom);
    QSpinBox *spinBox=new QSpinBox(&main);
    spinBox->setButtonSymbols(QAbstractSpinBox::NoButtons);

    progress=new QProgressBar(&main);
    progress->setMinimum(0);
    spinBox->setMaximum(300);
    QPushButton *startButton=new QPushButton(&main);
     stopButton=new QPushButton(&main);

    stopButton->setText("Стоп");
    stopButton->setEnabled(false);
    startButton->setText("Старт!");

    layoutH->addWidget(stopButton);
    layoutH->addWidget(startButton);

    layoutV->addWidget(label);

    layoutV->addWidget(spinBox);

    layoutV->addWidget(progress);
    layoutV->addLayout(layoutH);

    main.setLayout(layoutV);

    connect(spinBox,&QSpinBox::valueChanged,[&](int val){
        countWork=val;
        progress->setMaximum(val);
    });

    connect(startButton,&QPushButton::clicked,&main,[&](){

        stopButton->setEnabled(true);

         timer=new QTimer(this);

        connect(stopButton,&QPushButton::clicked,[&](){
             timer->stop();
             stopButton->setEnabled(false);

             progress->setValue(0);

             systemIcon.setIcon(QIcon(":/icon/icon/iconExcel.png"));
             systemIcon.show();

             QSystemTrayIcon::MessageIcon icon=QSystemTrayIcon::Information;
             QString title("!! Excel !!");
             QString text("Операция Остановилось!");
             systemIcon.showMessage(title,text,icon);


             message=new QMessageBox();
             message->addButton(QMessageBox::Ok);
             message->setText("Операция Остановилось!");
             message->exec();
         });

        listF=getFilesFolder();
        connect(timer,&QTimer::timeout,this,&MainWindow::Timer_Slot);

        timer->setInterval(1000);
        timer->start();
    });

    main.show();
  }
}

void MainWindow::Timer_Slot()
{

    on_buttonChek_clicked();
    on_buttonSave_clicked();

    countWork--;

    progress->setValue(progress->maximum()-countWork);

    if(countWork==0){
        timer->stop();

        progress->setValue(0);

        systemIcon.setIcon(QIcon(":/icon/icon/iconExcel.png"));
        systemIcon.show();

        QSystemTrayIcon::MessageIcon icon=QSystemTrayIcon::Information;
        QString title("!! Excel !!");
        QString text("Операция Успешно Выполнено!");
        systemIcon.showMessage(title,text,icon);

        message=new QMessageBox();
        message->addButton(QMessageBox::Ok);
        message->setText("Операция Успешно Выполнено!");
        message->exec();
    }
}

void MainWindow::Timer_Slot_Stop()
{
    timer->stop();
    stopButton->setEnabled(false);

    progress->setValue(0);

    systemIcon.setIcon(QIcon(":/icon/icon/iconExcel.png"));
    systemIcon.show();

    QSystemTrayIcon::MessageIcon icon=QSystemTrayIcon::Information;
    QString title("!! Excel !!");
    QString text("Операция Остановилось!");
    systemIcon.showMessage(title,text,icon);
    message=new QMessageBox();
    message->addButton(QMessageBox::Ok);
    message->setText("Операция Остановилось!");
    message->exec();
}

void MainWindow::on_checkBox_clicked(bool checked)
{
    newExcel->setVisible(checked);
}

void MainWindow::currentTimeBoxInit()
{
    if(currentSheet==2) ui->currentTimeBox->setCurrentIndex(0);
    if(currentSheet==3) ui->currentTimeBox->setCurrentIndex(1);
    if(currentSheet==4) ui->currentTimeBox->setCurrentIndex(2);
    if(currentSheet==5) ui->currentTimeBox->setCurrentIndex(3);
}

void MainWindow::setCurrentTimeBox(int crSheet)
{
    if(crSheet==0)currentSheet=2;
    if(crSheet==1)currentSheet=3;
    if(crSheet==2)currentSheet=4;
    if(crSheet==3)currentSheet=5;
}
