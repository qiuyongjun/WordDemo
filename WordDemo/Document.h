#ifndef DOCUMENT_H
#define DOCUMENT_H

#include <QObject>
#include <QAxObject>

#define NAMESPACE_BEGIN {
#define NAMESPACE_END }

namespace Word_NS
NAMESPACE_BEGIN

enum TitleLevel
{
    Level1 = 1, // 一级标题
    Level2, // 二级标题
    Level3, // 三级标题
    Level4, // 四级标题
};

class Document
{
public:
    Document();
	~Document();

public:
    // 初始化Word程序,耗时数秒，建议在非主线程调用
    static bool initWord();
    // 设置Word程序是否可见
    static bool setWordVisibel(bool visible);
    // 退出Word程序，不会保存更改
    static bool quitWord();

public:
    // 添加文本
    bool appendText(const QString & text);
    // 添加段落文本，段落文本结束会进行换行
    bool appendParagraphText(const QString & text);
    // 添加空段落，换行
    bool appendParagraph();
    // 添加指定大纲等级的标题
    bool appendTitle(const QString & text,TitleLevel level);
    // 插入图片
    bool insertPicture(const QString & filePath, const QString & pictureName = "");
    // 插入表格
    bool insertTable(int numRows, int numCols, const QStringList & tableData,const QString & tableName = "");

public:
	// 另存为
	bool saveAs(const QString & filePath);
	// 关闭文档
	bool close();

public:
    // 获取标题1样式，获取后可对其进行修改
    QAxObject *getTitle1Style() const;
    // 获取标题2样式，获取后可对其进行修改
    QAxObject *getTitle2Style() const;
    // 获取标题3样式，获取后可对其进行修改
    QAxObject *getTitle3Style() const;
    // 获取标题4样式，获取后可对其进行修改
    QAxObject *getTitle4Style() const;
    // 获取正文样式，获取后可对其进行修改
    QAxObject *getTextStyle() const;
    // 获取正文缩进样式，获取后可对其进行修改
    QAxObject *getTextIndent2Style() const;

private:
	// 初始化自定义样式
    bool initCustomStyle();

private:
    static QAxObject* m_word; // Word程序
    static bool m_visible; // Word程序是否可见，默认为不可见

private:
    QAxObject * m_document = nullptr; // 当前文档

    QAxObject * m_title1Style = nullptr; // 标题1样式
    QAxObject * m_title2Style = nullptr; // 标题2样式
    QAxObject * m_title3Style = nullptr; // 标题3样式
    QAxObject * m_title4Style = nullptr; // 标题4样式

    QAxObject * m_titleMultiLevelList = nullptr; // 标题多级列表

    QAxObject * m_textStyle = nullptr; // 正文

    QAxObject * m_textIndent2Style = nullptr; // 正文+首行缩进2字符

private:
    // 自动回收内存
    class Recycler
    {
    public:
        ~Recycler()
        {
            quitWord();
        }
    };
    static Recycler recycler;
};

NAMESPACE_END

#endif // DOCUMENT_H
