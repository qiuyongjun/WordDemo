#include "Document.h"

#include <QDebug>

#include <Windows.h>

namespace Word_NS
NAMESPACE_BEGIN


// 静态成员初始化
QAxObject* Document::m_word = nullptr;
// Word程序是否可见，默认为不可见
bool Document::m_visible = false;


Document::Document()
{
    // 如果未初始化Word程序，先进行初始化
    if(!m_word) initWord();

    // 获取Documents对象
    QAxObject *documents = m_word->querySubObject("Documents");

    // 添加一个新的文档
    m_document = documents->querySubObject("Add()");

    // 初始化样式
    initCustomStyle();
}

Word_NS::Document::~Document()
{
    close();
}

bool Document::initWord()
{
    if (m_word) return true;

    // 跨线程调用 Com
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    // 初始化Word程序
    m_word = new QAxObject("Word.Application");

    // 设置Word程序是否可见
    setWordVisibel(m_visible);

    return true;
}

bool Document::setWordVisibel(bool visible)
{
    m_visible = visible;

    // 程序未初始化
    if(!m_word) return true;

    // 设置Word程序是否可见
    bool result = m_word->setProperty("Visible", m_visible);

    return result;
}

bool Document::quitWord()
{
    // 程序未初始化
    if(!m_word) return true;

    // 对应CoInitializeEx
    OleUninitialize();

    // 退出Word程序,不保存更改
    m_word->dynamicCall("Quit(QVariant)", 0);
    // 清理Word
    m_word->clear();

    // 释放
    delete m_word;

    return true;
}

bool Document::appendText(const QString &text)
{
    if (!m_document) return false;

    // 激活当前文档，使之成为活动文档
    m_document->dynamicCall("Activate()");

    // 获取活动文档的当前选择范围
    QAxObject* selection  = m_word->querySubObject("Selection");

    // 插入文本
    selection->dynamicCall("TypeText(const QString&)",text);

    return true;
}

bool Document::appendParagraphText(const QString &text)
{
    if(!m_document) return false;

    // 激活当前文档，使之成为活动文档
    m_document->dynamicCall("Activate()");

    // 获取活动文档的当前选择范围
    QAxObject* selection  = m_word->querySubObject("Selection");

    // 插入文本
    selection->dynamicCall("TypeText(const QString&)",text);

    // 修改文本格式
    selection->dynamicCall("setStyle(WdBuiltinStyle)", m_textIndent2Style->asVariant());

    // 插入空段落
    selection->dynamicCall("TypeParagraph(void)");

    return true;
}

bool Document::appendParagraph()
{
    if (!m_document) return false;

    // 激活当前文档，使之成为活动文档
    m_document->dynamicCall("Activate()");

    // 获取活动文档的当前选择范围
    QAxObject* selection  = m_word->querySubObject("Selection");

    // 插入空段落
    selection->dynamicCall("TypeParagraph(void)");

    // 修改文本格式
    selection->dynamicCall("setStyle(WdBuiltinStyle)", m_textStyle->asVariant());

    return true;
}

bool Document::appendTitle(const QString &text, TitleLevel level)
{
    if(!m_document) return false;

    // 激活当前文档，使之成为活动文档
    m_document->dynamicCall("Activate()");

    // 获取活动文档的当前选择范围
    QAxObject* selection  = m_word->querySubObject("Selection");

    // 插入文本
    selection->dynamicCall("TypeText(const QString&)",text);

    //判断标题等级
    QAxObject * titleStyle = nullptr;
    switch (level)
    {
    case Level1:
    {
        titleStyle = m_title1Style;
        break;
    }
    case Level2:
    {
        titleStyle = m_title2Style;
        break;
    }
    case Level3:
    {
        titleStyle = m_title3Style;
        break;
    }
    case Level4:
    {
        titleStyle = m_title4Style;
        break;
    }
    }

    // 修改文本格式
    selection->dynamicCall("setStyle(WdBuiltinStyle)", titleStyle->asVariant());

    // 应用多级列表
    QAxObject * range = selection->querySubObject("Range");
    QAxObject * listFormat = range->querySubObject("ListFormat");
    listFormat->dynamicCall("ApplyListTemplate(QVariant)", m_titleMultiLevelList->asVariant());
    listFormat->setProperty("ListLevelNumber ",level);

    // 插入空段落
    selection->dynamicCall("TypeParagraph(void)");

    return true;
}

bool Word_NS::Document::insertPicture(const QString & filePath, const QString & pictureName)
{
    if (!m_document) return false;

    // 激活当前文档，使之成为活动文档
    m_document->dynamicCall("Activate()");

    // 获取活动文档的当前选择范围
    QAxObject* selection = m_word->querySubObject("Selection");

    // 修改文本格式
    selection->dynamicCall("setStyle(WdBuiltinStyle)", m_textStyle->asVariant());

    // 插入图片
    QAxObject* inlineShape = selection->querySubObject("InlineShapes")->querySubObject("AddPicture(QString)", filePath);

    // 将 InlineShape 转换为 Shape
    QAxObject *shape = inlineShape->querySubObject("ConvertToShape()");

    // 获取 WrapFormat 对象
    shape->querySubObject("WrapFormat")->setProperty("Type", 4); // 上下型环绕

    // 获取 Range
    QAxObject * range = selection->querySubObject("Range");

    // 插入图片标题
    if (!pictureName.isEmpty())
    {
        QAxObject *captionLabels = m_word->querySubObject("CaptionLabels");
        QAxObject *captionLabel = captionLabels->querySubObject("Item(Variant)", QString::fromLocal8Bit("图")); // “图” 标签

        // 插入题注
        QVariantList params;
        params << captionLabel->asVariant() << pictureName << QVariant() << QVariant(1) << QVariant();	// 在图片下方插入题注
        range->dynamicCall("InsertCaption(QVariant, QVariant, QVariant, QVariant, QVariant)", params);
    }

    // 更新 selection 对象，跳出题注
    selection->dynamicCall("Move(QVariant, QVariant)", 4, 1); // 4 表示 wdParagraph，1 表示移动到下一个段落的开头
    selection->dynamicCall("TypeParagraph(void)"); // 插入空段落

    return true;
}

bool Document::insertTable(int numRows, int numCols, const QStringList &tableData, const QString &tableName)
{
    if (!m_document) return false;

    // 激活当前文档，使之成为活动文档
    m_document->dynamicCall("Activate()");

    // 获取活动文档的当前选择范围
    QAxObject* selection = m_word->querySubObject("Selection");

    // 修改文本格式
    selection->dynamicCall("setStyle(WdBuiltinStyle)", m_textStyle->asVariant());

    // 插入表格
    QAxObject * range = selection->querySubObject("Range");
    QAxObject * tables = m_document->querySubObject("Tables");
    QAxObject * table = tables->querySubObject("Add(QVariant, QVariant, QVariant, QVariant, QVariant, QVariant)", range->asVariant(), numRows, numCols, 1, 1);


    // 填充表格内容
    int dataIndex = 0; // 用于从 tableData 中获取数据的索引
    for (int row = 1; row <= numRows; ++row)
    {
        if (dataIndex >= tableData.size()) break;

        for (int col = 1; col <= numCols; ++col)
        {
            QAxObject *cell = table->querySubObject("Cell(int,int)", row, col);
            QAxObject *cellRange = cell->querySubObject("Range");

            cellRange->dynamicCall("InsertAfter(QString)", tableData.at(dataIndex));
            ++dataIndex;

        }
    }

    // 获取 tableRange
    QAxObject *tableRange = table->querySubObject("Range");

    // 插入表格标题
    if (!tableName.isEmpty())
    {
        QAxObject *captionLabels = m_word->querySubObject("CaptionLabels");
        QAxObject *captionLabel = captionLabels->querySubObject("Item(Variant)", QString::fromLocal8Bit("表格")); // “表格” 标签

        // 插入题注
        QVariantList params;
        params << captionLabel->asVariant() << tableName << QVariant() << QVariant(0) << QVariant(); // 在表格上方插入题注
        tableRange->dynamicCall("InsertCaption(QVariant, QVariant, QVariant, QVariant, QVariant)", params);
    }

    // 根据内容自适应宽度
    table->dynamicCall("AutoFitBehavior(QVariant)", 1);
    // 根据窗口自适应宽度
    table->dynamicCall("AutoFitBehavior(QVariant)", 2);

    // 更新 selection 对象，跳出表格
    auto tableRangeEnd = tableRange->property("End");
    selection->dynamicCall("SetRange(QVariant, QVariant)", tableRangeEnd, tableRangeEnd);

    return true;
}

bool Word_NS::Document::saveAs(const QString & filePath)
{
    if (!m_document) return false;

    // 另存为
    m_document->dynamicCall("SaveAs(const QString&)", filePath);

    return true;
}

bool Document::initCustomStyle()
{
    // 使用默认ListGallery（编号列表）
    QAxObject *listGallerys = m_word->querySubObject("ListGalleries");
    QAxObject *listGallery = listGallerys->querySubObject("Item(QVariant)", 3); // 3 表示多级列表

    // 获取默认ListTemplate
    QAxObject *listTemplates = listGallery->querySubObject("ListTemplates");
    m_titleMultiLevelList = listTemplates->querySubObject("Item(QVariant)", 2); // 2 表示模板第二个，不包括无

    // 创建自定义样式
    QAxObject *styles = m_document->querySubObject("Styles");

    // 添加Title1样式
    m_title1Style = styles->querySubObject("Add(QVariant,QVariant)", "Title1",1); // wdStyleTypeParagraph	1	段落样式
    m_title1Style->querySubObject("Font")->setProperty("Name",QString::fromLocal8Bit("宋体"));
    m_title1Style->querySubObject("Font")->setProperty("Size",12); // 小四
    m_title1Style->querySubObject("ParagraphFormat")->setProperty("OutlineLevel",1); // 大纲级别 1
    m_title1Style->querySubObject("ParagraphFormat")->setProperty("LineSpacingRule", 1); // wdLineSpace1pt5 1  1.5倍行距

    // 添加Title2样式
    m_title2Style = styles->querySubObject("Add(QVariant,QVariant)", "Title2",1); // wdStyleTypeParagraph	1	段落样式
    m_title2Style->querySubObject("Font")->setProperty("Name",QString::fromLocal8Bit("宋体"));
    m_title2Style->querySubObject("Font")->setProperty("Size",12);
    m_title2Style->querySubObject("ParagraphFormat")->setProperty("OutlineLevel",2); // 大纲级别 2
    m_title2Style->querySubObject("ParagraphFormat")->setProperty("LineSpacingRule", 1); // wdLineSpace1pt5 1  1.5倍行距

    // 添加Title3样式
    m_title3Style = styles->querySubObject("Add(QVariant,QVariant)", "Title3",1); // wdStyleTypeParagraph	1	段落样式
    m_title3Style->querySubObject("Font")->setProperty("Name",QString::fromLocal8Bit("宋体"));
    m_title3Style->querySubObject("Font")->setProperty("Size",12);
    m_title3Style->querySubObject("ParagraphFormat")->setProperty("OutlineLevel",3); // 大纲级别 3
    m_title3Style->querySubObject("ParagraphFormat")->setProperty("LineSpacingRule", 1); // wdLineSpace1pt5 1  1.5倍行距

    // 添加Title4样式
    m_title4Style = styles->querySubObject("Add(QVariant,QVariant)", "Title4",1); // wdStyleTypeParagraph	1	段落样式
    m_title4Style->querySubObject("Font")->setProperty("Name",QString::fromLocal8Bit("宋体"));
    m_title4Style->querySubObject("Font")->setProperty("Size",12);
    m_title4Style->querySubObject("ParagraphFormat")->setProperty("OutlineLevel",4); // 大纲级别 4
    m_title4Style->querySubObject("ParagraphFormat")->setProperty("LineSpacingRule", 1); // wdLineSpace1pt5 1  1.5倍行距

    // 添加Text样式
    m_textStyle = styles->querySubObject("Add(QVariant,QVariant)", "Text", 1); // wdStyleTypeParagraph	1	段落样式
    m_textStyle->querySubObject("Font")->setProperty("Name", QString::fromLocal8Bit("宋体"));
    m_textStyle->querySubObject("Font")->setProperty("Size", 12);
    m_textStyle->querySubObject("ParagraphFormat")->setProperty("Alignment", 0); // 0 左对齐
    m_textStyle->querySubObject("ParagraphFormat")->setProperty("LineSpacingRule", 1); // wdLineSpace1pt5 1  1.5倍行距

    // 添加TextIndent2样式
    m_textIndent2Style = styles->querySubObject("Add(QVariant,QVariant)", "TextIndent2",1); // wdStyleTypeParagraph	1	段落样式
    m_textIndent2Style->querySubObject("Font")->setProperty("Name",QString::fromLocal8Bit("宋体"));
    m_textIndent2Style->querySubObject("Font")->setProperty("Size",12);
    m_textIndent2Style->querySubObject("ParagraphFormat")->setProperty("CharacterUnitFirstLineIndent", 2); // 首行缩进2字符
    m_textIndent2Style->querySubObject("ParagraphFormat")->setProperty("LineSpacingRule", 1); // wdLineSpace1pt5 1  1.5倍行距
    m_textIndent2Style->querySubObject("ParagraphFormat")->setProperty("Alignment", 0); // 0 左对齐

    // 修改默认题注样式
    QAxObject * captionStyle = styles->querySubObject("Item(QVariant)", QString::fromLocal8Bit("题注"));
    captionStyle->querySubObject("Font")->setProperty("Name", QString::fromLocal8Bit("宋体"));
    captionStyle->querySubObject("Font")->setProperty("Size", 12); // 小四
    captionStyle->querySubObject("ParagraphFormat")->setProperty("LineSpacingRule", 1); // wdLineSpace1pt5 1  1.5倍行距
    captionStyle->querySubObject("ParagraphFormat")->setProperty("Alignment", 1); // 1 表示居中对齐

    return true;
}

bool Word_NS::Document::close()
{
    if (!m_document) return false;

    // 关闭文档，不会保存
    m_document->dynamicCall("Close()");

    //
    m_document = nullptr;

    return true;
}

QAxObject *Document::getTitle1Style() const
{
    return m_title1Style;
}

QAxObject *Document::getTitle2Style() const
{
    return m_title2Style;
}

QAxObject *Document::getTitle3Style() const
{
    return m_title3Style;
}

QAxObject *Document::getTitle4Style() const
{
    return m_title4Style;
}

QAxObject *Document::getTextStyle() const
{
    return m_textStyle;
}

QAxObject *Document::getTextIndent2Style() const
{
    return  m_textIndent2Style;
}



NAMESPACE_END
