#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include"excel.h"
#include <QMainWindow>
#include<QtWidgets>
#include<QFileDialog>
#include<QSettings>
#include<QAxObject>
#include<QFileInfo>
#include<QSystemTrayIcon>
#include<algorithm>
QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE
class MainWindow : public QMainWindow
{
    Q_OBJECT
public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
private slots:
    void on_buttonChooseOldPath_clicked();
    void on_buttonNewExcelFile_clicked();
    void on_buttonChek_clicked();
    void on_buttonSave_clicked();
    void on_pushButton_clicked();
    void on_pushButton_2_clicked();
    void Timer_Slot();
    void Timer_Slot_Stop();
    void on_checkBox_clicked(bool checked);

private:
    void currentTimeBoxInit();
    void setCurrentTimeBox(int crSheet);
    void extracted(QVector<QVariant> &rowsV);
    void buttonsEnabled(bool is);
    QVector<QLineEdit*> lineEdist;
    QVector<QLineEdit*> lineTest;
    void getLists();
    QVector<QLineEdit*> lineLastEdits;
    Ui::MainWindow *ui;
    QSettings *settings;
    QLabel *fileFolderStatus;
    QLabel *excelFileStatus;
    unsigned int day=1;
    unsigned int currentSheet=2;
    unsigned const int kolSheet=5;
    QLabel *label_currentFileOpen;
    QStringList getFilesFolder();
    QStringList month;
    Excel *excel;
    Excel *newExcel;
    int currentFileIndex=0;
    int row;
    bool isNewFileOpen;
    int countWork=2;
    QMessageBox *message;
    QTimer *timer;
    QProgressBar *progress;
    QSystemTrayIcon systemIcon;
    QPushButton *stopButton;
    QStringList  listF;
    QWidget main;
};
#endif // MAINWINDOW_H
