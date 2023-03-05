#ifndef WPSCOMMONAPI_H
#define WPSCOMMONAPI_H

#include <QDebug>
#include <QHBoxLayout>
#include <QMap>
#include <QMutex>
#include <QQueue>
#include <QX11Info>
#include <QKeyEvent>
#include <QFocusEvent>
#include <QLineEdit>
#include <QTextEdit>

//wpsapi.h 头文件引用，必须要kfc/comsptr.h , ksoapi/ksoapi.h , wpsapiex.h 引用之前
#include "wpsapi.h"

#include "kfc/comsptr.h"
#include "ksoapi/ksoapi.h"
#include "wpsapiex.h"

#if (QT_VERSION <= QT_VERSION_CHECK(5, 0, 0))
#include <QtGui/QMenuBar>
#include <QtGui/QX11EmbedContainer>
#else
#include <QWidget>
#endif
#include "MemoryPool.h"
#include "rightclickmodel.h"
#include "wpsdatahelper.h"
#include <vector>


#include "ksocontainerwidget.h"

using namespace wpsapi;
using namespace kfc;
using namespace ksoapi;
using namespace std;

#define WPSModule "WPS"

typedef HRESULT (*callBackFun)(QVariantList& args);

#if (QT_VERSION < QT_VERSION_CHECK(5, 0, 0))
class WpsCommonAPI : public QX11EmbedContainer
#else
class WpsCommonAPI : public KSOContainerWidget,
                     public QAbstractNativeEventFilter
#endif
{
    Q_OBJECT
public:
    explicit WpsCommonAPI(QWidget* parent = nullptr);

    ~WpsCommonAPI();

protected:
    bool eventFilter(QObject *Object, QEvent *Event);
#if (QT_VERSION >= QT_VERSION_CHECK(5, 0, 0))
    bool nativeEventFilter(const QByteArray &eventType, void *message, long *result);
#endif

public:
    void SendFocus();

    bool InitWpsApplication();

    void DestroyWpsApplication();

    bool initWps();

    bool getSelectFontfamily(QString& fontFamily);

    bool getSelectFontSize(float& fontSize);

    bool newDocument();

    //页眉的图标和文字
    bool setHeaderFooterInfo(const QString& picPath, int picWidth, int picHeight, const QString& title, bool bHeader = true);

    bool addPicture(ks_stdptr<Range> ptr_range, const QString& picPath, int picWidth, int picHeight, bool bWatermark = true);

    bool getPageRange(ks_stdptr<_Document> ptr_document, int pageIndex, ks_stdptr<Range>& ptr_range);

    bool addPictureToPages(const QString& picPath, int picWidth, int picHeight, const QList<int>& pageList);

    long getTotalPages();

    bool excuteCombars(const QString& key, const QString& key2);

    //获取鼠标所在的range的结束位置长度为0
    //选择的文字的字体
    bool getSelectionFont(ks_stdptr<_Font>& font);
    //命令工具栏
    bool getCommandBars(ks_stdptr<_CommandBars>& cmdBars);
    //获取doc
    bool getDocument(ks_stdptr<_Document>& _document);

    bool getApplication(ks_stdptr<_Application>& app);
    //获取选择的range
    bool getSelectionRange(ks_stdptr<Range>& range);

    bool getSelectionStartRange(ks_stdptr<Range>& range);
    //获取选择列表格式
    bool getSelectionListFormat(ListFormat** listForamt);
    //获取选择的
    bool getSelection(ks_stdptr<Selection>& _selection);

    //获取选择的段落
    bool getSelectionParagraphFormat(ParagraphFormat** paragraphFormat);

    //执行一些MSO指令，调用wps的按钮命令
    bool ExecuteMso(BSTR mscID);

    IKRpcClient* m_pRpcClient = nullptr;

    //隐藏菜单
    bool hiddenHeaderMenu();

    long getContentLength();

    //设置字体
    bool setFontfamily(const QString& family);

    //设置字体大小ks_stdptr<Range>
    bool setFontSize(float size);

    //设置粗体
    bool setFontBold();

    //设置斜体
    bool setFontItalic();

    //设置xiahuaxian
    bool setFontUnderline();

    bool setFontDoubleUnderline();

    bool setFontDeleteline();

    bool setFontWithTop();

    bool setFontWithBottom();

    bool enlargFont();

    bool reduceFont();

    bool Indent(bool enlarge);

    bool Undo();

    bool restore();

    bool copyCmd();

    bool cutCmd();

    bool pasteCmd();

    //圆点列表std
    bool setListType_mutilLevel();

    //打开字体窗口
    bool showFontDialog();

    //页面设置界面
    bool showPageSetup();

    bool canSetPageSetUp();

    bool showBookMark();

    //设置左对其
    bool setLeftAlignment();

    //设置居中
    bool setCenterAlignment();

    //设置右对齐
    bool setRightAlignment();

    //设置
    bool setFullAlignment();

    bool setLineSpacing(float spacing);

    //圆点列表std
    bool setListType_point();

    //数字列表 void DestroyWpsApplication();
    bool setListType_number();

    //格式功能
    bool setFormatPainter(bool isSet);

    bool setHeader();

    bool setFooter();

    bool setPageNumber(bool top);

    bool showFind();

    bool showSave();

    bool showReplace();

    //打印
    bool showPrint();

    bool showPrintPreview();
    //开启转写
    bool startTransform(bool isEnd);

    //停止转写
    bool stopTransform(bool isStopByStart = false);

    //添加临时文件
    bool addTmpText(const QString& message);

    //添加最终转写文字
    bool addFinnalText(const QString& message, int bookMark);

    //在选择的位置添加文本内容
    bool addTextAtSelection(const QString& message, bool isConver = false);

    bool deleteAfterIfNoSelectText();

    //保存wps
    bool saveWps(const QString& fileName);

    bool saveWps();

    bool openFile(const QString& fileName, bool isreadOnly = false);

    //弹出打印界面
    bool printWps();

    bool getSelectBookMark(QString& strStart, QString& strEnd);

    bool nextParagraph();

    bool addMenus();

    bool setNowTime(long target, bool isFollow = false);

    bool init_GetStartTime(long& time);

    bool initBookMarks();

    bool focusOnTransform();

    bool resetListenBackArea(bool isCacheClear);

    //标记区域
    bool resetMarkArea(bool isCacheClear);

    bool restoreMarkArea();

    bool getWordEndRange(ks_stdptr<Range>& range);

    // 有无内容
    bool hasContent();

    bool getSelectTime(long& strStart);

    bool getSelectTime(long& strStart, ks_stdptr<Range>& ptr_range);

    bool getBookMarkByName(QString name, ks_stdptr<Bookmark>& bookmark);

    bool initWpsData();

    bool getSelectText(QString& text, long maxLength = -1, bool removeSpace = false);

    bool addMicName(const QString& micName);

    bool setFullScreen(bool isSet = true);

    QMutex _mutex;
    QMutex _mutexForTmp;

    CMemoryPool<rightClickModel> m_rcmPool;

    void tryEnqueue(rightClickModel* msg);
    bool tryDequeue(rightClickModel& msg);

    bool EnableStartInsert(bool enable = true);
    bool EnableAddHotword(bool enable = true);
    bool EnableListenBack(bool enable = true);
    bool EnableMark(bool enable = true);
    bool EnableMarkCancle(bool enable = true);
    bool EnableStopInsert(bool enable = true);

    //对应的书签的颜色变换
    bool setMarkBackgroundByTime(long nowTime, bool ismark);

    QHash<int, ks_stdptr<Range>> qhashRanges;

    bool setSpanTime(long time);

    bool mark(bool isMark);

    bool setForeground(int red, int green, int black);

    bool setBackground(int red, int green, int black);

    bool MergeSameMicNameContent();
    bool myTest();

    void setTransformFontFamily(const QString& font);

    void setTransformFontSize(const float& fontSize);

    bool setTextMenuVisble(bool enable);
    bool removeStartPunctuation(QString& src);

    ///
    /// \brief AppendFocusObject 需要特殊处理的非wps控件，当焦点再这些控件里的时候，键盘输入事件仍旧送入wps，避免wps内部的中文输入法卡死
    /// \param obj 非wps控件（涉及到焦点）
    /// \return
    ///
    int AppendFocusObject(QObject *obj);

private slots:
    /**
    * @describe  WPS提供修复渲染问题
    */
    void onEmbeded();

private:
    QHBoxLayout* m_hlayout = nullptr;
    QWidget* m_containerWin = nullptr;

    QString m_transformFamily = "";

    single m_transformFontSize = 12;

    single m_maxFontSize = 72;

    float m_minFontSize = 5;

    QString m_pre_micName = "";

    bool m_hasContent = false;

    ks_stdptr<Range> ptrRangeTest;

    bool m_isTransforming = false;
    //老的转写信息

    //documentFiled属性，只读的，这个界面删除了其实还在，很神奇的东西
    ks_stdptr<DocumentField> m_rangeForMark;

    QQueue<rightClickModel*> queueRightMsg;

    QMutex _mutexQueue;

    void debugRangeText(ks_stdptr<Range> ptr_range);

    void debugRangeStartEnd(ks_stdptr<Range> ptr_range);

    bool setRangeBackground(ks_stdptr<Range>& ptr_range, WdColor color);

    // bool splitTwoRangeByMark(ks_stdptr<Range> &ptr_range, QString msg, ks_stdptr<Range> &ptr_range_first, ks_stdptr<Range> &ptr_range_end);

    QString m_markText = "【正在转写...】";

    QString m_markHead = "mark_ifly_";

    //时间戳集合
    vector<long> m_list_time;

    //range集合
    vector<ks_stdptr<Bookmark>> m_list_bookmark;

    bool addMenu(const QString& cmdName, const QString& caption, callBackFun fun, ks_stdptr<CommandBarControl>& ptrCtrl);

    bool addMenuTest(const QString& cmdBarName, const QString& caption);

    //初始化产生的应用 的指针对象
    ks_stdptr<_Application> m_spApplication;
    ks_stdptr<wpsapiex::_ApplicationEx> m_spApplicationEx;
    //app的doc
    ks_stdptr<Documents> m_docs;

    bool getMarkRangeByTmpRange(ks_stdptr<Range>& ptrRange, const QString& text);

    bool getRangeEndofRaneg(ks_stdptr<Range>& ptrRange, ks_stdptr<Range>& ptrGet);

    bool getRangeStartofRaneg(ks_stdptr<Range>& ptrRange, ks_stdptr<Range>& ptrGet);

    long getRangeEnd(ks_stdptr<Range>& ptrRange);
    long getRangeStart(ks_stdptr<Range>& ptrRange);

    bool resetTextRange(const ks_stdptr<Range>& ptrRange);

    ks_stdptr<Range> m_rangeForText;

    ks_stdptr<Range> m_range_micName;

    void setRange(ks_stdptr<Range>& ptrRange, ks_stdptr<Range>& ptrNewRange);

    QString getRangeText(ks_stdptr<Range>& ptrRange);

    bool setRangeText(ks_stdptr<Range>& ptrRange, QString text);

    void setNewDocumentField(ks_stdptr<DocumentField>& ptrDocField, ks_stdptr<DocumentField>& ptrNewDoc);
    //正在回听的书签，如果是相同的书签，不做处理，节约时间
    ks_stdptr<Bookmark> m_playing_bookmark;

    //正在回听的range,缓存用，下次切换了先变颜色，节约时间
    ks_stdptr<Range> m_playing_range;

    //初始化all的书签
    bool m_isInitialBookmarksOK = false;

    WdColor getColorByRGB(int r, int g, int b);

    //检查是否需要重新设置 转写mark
    bool resetMark();

    //二分法
    long binaryFindFirstMoreNum(vector<long>& nums, long n);

    //获取鼠标所在位置的第一个书签的时间消息
    static WpsCommonAPI* m_innerWps;

    //右键菜单
    static HRESULT rightclick_callBackMark(QVariantList& args);

    static HRESULT rightclick_callBackMarkCancle(QVariantList& args);

    static HRESULT rightclick_addHotWord(QVariantList& args);

    static HRESULT rightclick_listenBack(QVariantList& args);

    static HRESULT rightclick_stopInsertTransform(QVariantList& args);

    static HRESULT rightclick_startInsertTransform(QVariantList& args);

    //static HRESULT WpsOnSelectionCallback(QVariantList &args);

    //清除右健菜单
    bool CleanMenu(QString cmdName, vector<QString> listName, vector<QString> listDelete);

    ks_stdptr<Bookmarks> m_doc_bookmarks;

    bool getDocBookMarks(ks_stdptr<Bookmarks>& cmdBars);

    ks_stdptr<_Document> m_doc;

    QString bstrToStdQString(BSTR strData);

    //书签的间隔
    long m_mark_span = 300;

    //QDateTime qtime_lastFinal;

    long m_pre_playtime = -1;

    //通过当前的时间和设定的时间间隔 来获取 书签的值，511-》500
    long getSpanData(long data);

    QString timeToBookMakrName(long data);

    //上次的书签的名称
    QString preMarkText;

    //书签对应的rang
    ks_stdptr<Range> m_pre_bookmark_range;

    QString m_pre_final_msg = "";

    bool CanChangeStyle(long start, long end);

    bool repairMicName();

    ks_stdptr<CommandBarControl> ptr_startInsert;
    ks_stdptr<CommandBarControl> ptr_stopInsert;
    ks_stdptr<CommandBarControl> ptr_mackCancle;
    ks_stdptr<CommandBarControl> ptr_mark;
    ks_stdptr<CommandBarControl> ptr_listenBack;
    ks_stdptr<CommandBarControl> ptr_addHotWord;

    ks_stdptr<CommandBarControl> ptr_startInsert_Bulleted;
    ks_stdptr<CommandBarControl> ptr_stopInsert_Bulleted;
    ks_stdptr<CommandBarControl> ptr_mackCancle_Bulleted;
    ks_stdptr<CommandBarControl> ptr_mark_Bulleted;
    ks_stdptr<CommandBarControl> ptr_listenBack_Bulleted;
    ks_stdptr<CommandBarControl> ptr_addHotWord_Bulleted;

    ks_stdptr<CommandBarControl> ptr_startInsert_Numbered;
    ks_stdptr<CommandBarControl> ptr_stopInsert_Numbered;
    ks_stdptr<CommandBarControl> ptr_mackCancle_Numbered;
    ks_stdptr<CommandBarControl> ptr_mark_Numbered;
    ks_stdptr<CommandBarControl> ptr_listenBack_Numbered;
    ks_stdptr<CommandBarControl> ptr_addHotWord_Numbered;

    bool setRangeForeground(ks_stdptr<Range>& ptr_range, WdColor color);

    bool addBookMarkOnRange(ks_stdptr<Range>& range, QString& bookmakrName, ks_stdptr<Bookmark>& bookmark, bool isReturn = false, bool appenChar = false);

    QString getMicBookMarkName(const QString& name);

    QString getTextOnBookMark(ks_stdptr<Bookmark>& bookmark);

    QString getTextOnParagraph(ks_stdptr<Paragraph>& paragraph);

    bool getRangeOnParagraph(ks_stdptr<Paragraph>& paragraph, ks_stdptr<Range>& ange);

    bool getParagraphFirstMic(ks_stdptr<Paragraph>& paragraph, QString& micName, long& rangeStart, long& rangeEnd);

    bool setRangeFontFamily(ks_stdptr<Range>& ptrRange, QString text);
    bool setRangeFontSize(ks_stdptr<Range>& ptrRange, single size);
    bool getRangeByStartEnd(long start, long end, ks_stdptr<Range>& para_range);
    bool setRightMenuVisble(const QString& name, bool enable);
    bool changedSelection();
    bool ResetDocumentFieldBYRange(ks_stdptr<Range> ptr_range);
    QString getFieldText(ks_stdptr<DocumentField>& ptr_doc);

    void deleteCommandBarControl(ks_stdptr<CommandBarControl>& para_cmbar);

    WpSDataHelper wpsHelper;

    // 非wps内部的控件列表
    QList<QObject*> m_FocusWidget_lst;

    // 当前最新转写结果（最终）的时间戳
    unsigned int m_current_timestamp;

    // 是否完成了一次最终结果的插入
    bool m_addFinal_finished;
};

#endif // WPSCOMMONAPI_H
