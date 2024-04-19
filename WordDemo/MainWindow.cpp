#include "MainWindow.h"
#include "ui_MainWindow.h"

#include "Document.h"

#include <QDir>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    QString filePath = QCoreApplication::applicationDirPath() + "/document.docx";
    ui->lineEdit_filePath->setText(filePath);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_clicked()
{

    Word_NS::Document document;

    document.appendTitle(QString::fromLocal8Bit("标题1"), Word_NS::Level1);
    document.appendParagraphText("tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt");

    document.appendTitle(QString::fromLocal8Bit("标题2"), Word_NS::Level1);
    document.appendParagraphText("tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt");

    document.insertPicture(QCoreApplication::applicationDirPath() + "/theme.jpg", QString::fromLocal8Bit("功能架构"));
    document.appendParagraphText("tttttttttttttttttttttttttttttttttttttttttttttt"
        "tttttttttttttttttttttttttttttttttttttttttttttt"
        "tttttttttttttttttttttttttttttttttttttttttttttt"
        "tttttttttttttttttttttttttttttttttttttttttttttt"
        "tttttttttttttttttttttttttttttttttttttttttttttt"
        "tttttttttttttttttttttttttttttttttttttttttttttt");

    document.appendTitle(QString::fromLocal8Bit("标题3"), Word_NS::Level1);
    document.appendParagraphText("tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt");

	QStringList tableData =
	{
		"1","2","3",
		"4","5","6",
		"7","8","9"
	};
    document.insertTable(3, 3, tableData, QString::fromLocal8Bit("测试报告"));

    document.appendParagraphText("tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt"
		"tttttttttttttttttttttttttttttttttttttttttttttt");

    document.saveAs(ui->lineEdit_filePath->text());
}
