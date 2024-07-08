#ifndef EXCEL_H
#define EXCEL_H

#include<QtWidgets>
#include<QAxObject>
#include<QVector>
struct ExcelSize
{
    unsigned int RowStart;
    unsigned int ColStart;
    unsigned int cols;
    unsigned int row;
};

class Excel{
private:
    unsigned int sheet;
    QString pathFile;
    QAxObject *excel;
    QAxObject *workbooks;
    QAxObject *workbook;
    QAxObject *worksheet;
    ExcelSize size;
public:
    Excel(QString pathFile,unsigned int Sheet):sheet(Sheet),pathFile(pathFile)
    {
        for(int i=0;i<pathFile.size();++i)
        {
            if(pathFile[i]=='\\')
            {
                pathFile[i]='/';
            }
        }

        excel=new QAxObject("Excel.Application");
        workbooks=excel->querySubObject("Workbooks");
        workbook=workbooks->querySubObject("Open(const QString&)",pathFile);
        worksheet = workbook->querySubObject("WorkSheets(int)",sheet);

    }

    void setVisible(bool b){
         excel->dynamicCall("SetVisible(bool)",b);
    }


    void setSheet(int n){
            worksheet = workbook->querySubObject("WorkSheets(int)",n);
    }
    void setPathFile(QString path)
    {
            workbook=workbooks->querySubObject("Open(const QString&)",path);
    }

    QVariant WriteData(unsigned int i,unsigned int j)
    {
        QAxObject *Cell=worksheet->querySubObject("Cells(int,int)",i,j);
        return Cell->dynamicCall("Value()");
    }

    QVariant WriteData(QString point)
    {
        QAxObject *Cell=worksheet->querySubObject("Range(const QString &)",point);
        return Cell->dynamicCall("Value()");
    }

    void setData(int i,int j,QVariant data){
        QAxObject *cell = worksheet->querySubObject("Cells(int, int)", i, j);
        cell->setProperty("Value", data.toString());
    }

    void setData(QString point,QVariant data)
    {
        QAxObject *Cell=worksheet->querySubObject("Range(const QString &)",point);
        Cell->setProperty("Value",data.toString());
    }

    ExcelSize getExcelSize(){
        ExcelSize temp;

        QAxObject *userange=worksheet->querySubObject("UsedRange");
        QAxObject *columns = userange->querySubObject("Columns");
        QAxObject *rows=userange->querySubObject("Rows");

        temp.RowStart=userange->property("Rows").toInt();
        temp.ColStart=userange->property("Column").toInt();
        temp.cols=columns->property("Count").toInt();
        temp.row=rows->property("Count").toInt();

        return temp;
    }

    void deleteRow(int row){
        worksheet->querySubObject("Rows(int)", row)->dynamicCall("Delete()");
    }

    QVector<QVariant> getColVector(int col)
    {
        QVector<QVariant> result;
        ExcelSize size=getExcelSize();

        for(int i=1;i<size.row;++i)
        {
            result.push_back(this->WriteData(i,col).toString());
        }

        return result;
    }

    QVector<QVariant> getColVector(QString col)
    {
        QVector<QVariant> result;

        ExcelSize size=getExcelSize();

        for(int i=1;i<size.row;++i)
        {
            result.push_back(WriteData(col+QString::number(i)));
        }

        return result;
    }

    QVector<QVariant> getRowVector(int row)
    {
        QVector<QVariant> result;
        ExcelSize size=getExcelSize();

        for(int i=1;i<size.cols;++i)
        {
            result.push_back(this->WriteData(row,i).toString());
        }

        return result;
    }

    QString getNameSheet(){
        QString name=this->worksheet->property("Name").toString();
        return name;
    }

    int getCountSheet()
    {
        qDebug()<<"Мы Тут";
        QAxObject *sh=workbook->querySubObject("WorkSheets");
        int count=sh->property("Count").toInt();
        return count;
    }

    void setTableWidget(QTableWidget &table)
    {
        size=getExcelSize();

        table.setColumnCount(size.ColStart+size.cols);
        table.setRowCount(size.RowStart+size.row);

        for(int i=0;i<size.row;++i)
        {
            for(int j=0;j<size.cols;++j)
            {
                QVariant temp = WriteData(i+1,j+1);
                QTableWidgetItem *item=new QTableWidgetItem(temp.toString());
                table.setItem(i,j,item);
            }
        }
    }

        QVector<QVariant> getLastTenValues(QString col,int row){

            QVector<QVariant> result;

            int k=10;
            if(row<10){
                k=row;
            }
                for(int i=row;i>row-k;--i){
                    result.push_back(this->WriteData(col+QString::number(i)));
                }
            return result;
         }
    void saveExcel()
    {
        workbook->dynamicCall("Save()",true);
    }
    void close()
    {
        workbook->dynamicCall("Close");
        excel->dynamicCall("Quit()");
    }
};
#endif // EXCEL_H
