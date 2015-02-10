#include <QtTest>
#include <QString>
#include <QDebug>
#include <QByteArray>
#include <QFile>
#include <QFont>
#include <QTime>
#include <QColor>
#include <QDomDocument>

#include <QImage>
#include <QMimeDatabase>
#include <QMimeType>


#include <docx/document.h>
#include <docx/text.h>
#include <docx/table.h>


using namespace Docx;

class TestDocument : public QObject
{
    Q_OBJECT

public:
    TestDocument();

private Q_SLOTS:
    void testLoad();
    void testImageInfo();
    void testTable();
};


TestDocument::TestDocument()
{

}

const QString imagePath1("://c.png");
const QString imagePath2("://139924.jpg");
const QString imagePath3("://149341.jpg");
const QString imagePath4("://185607.jpg");

void TestDocument::testLoad()
{
    //://Full2.docx
    //://Full.docx
    //://default.docx

    Document doc;
    doc.addHeading("MyTitle", 0);
    Paragraph *p = doc.addParagraph("helleWord");
    p->insertParagraphBefore("Before", "ListBullet");
    p->insertParagraphBefore("Number1", "ListNumber");
    p->insertParagraphBefore("Number2", "ListNumber");

    Run *addRun = p->addRun("AddRun");
    addRun->setStyle(QStringLiteral("Emphasis"));
    addRun->addText("Hello");

    doc.addParagraph();
    doc.addHeading("MyHead1");

    Paragraph *emptyP = doc.addParagraph();
    Run * tabRun = emptyP->addRun();
    tabRun->addTab();
    emptyP->addRun("EmptyParagraph", "IntenseEmphasis");

    doc.addParagraph("West", "IntenseQuoteChar");

    Paragraph *p2 = doc.addParagraph("Two");
    Run *r2 = p2->addRun("Next");
    r2->setBold(true);
    r2->setBold(false);
    r2->setBold(true);

    doc.addHeading("Heading2", 2);
    Paragraph *p3 = doc.addParagraph();
    Run *run = p3->addRun("MyContent");
    run->addText("abc");
    run->setBold(true);
    run->setUnderLine(WD_UNDERLINE::SINGLE);
    run->addTab();

    run = p3->addRun("Main");
    run->setAllcaps(true);
    run->setUnderLine(WD_UNDERLINE::DASH);
    run->addTab();

    run = p3->addRun("Main2");
    run->setDoubleStrike(true);
    run->setUnderLine(WD_UNDERLINE::DASH_HEAVY);
    run->addTab();

    run = p3->addRun("Main3");
    run->setItalic(true);
    run->setUnderLine(WD_UNDERLINE::DOT_DASH);
    run->addTab();

    run = p3->addRun("Main3");
    run->setSmallcaps(true);
    run->setUnderLine(WD_UNDERLINE::None);
    run->setShadow();
    run->addTab();


    qDebug() << p3->text();

    doc.addParagraph();
    p2 = doc.addParagraph("Alignment1");
    p2->setAlignment(WD_PARAGRAPH_ALIGNMENT::CENTER);

    p2 = doc.addParagraph("Alignment2");
    p2->setAlignment(WD_PARAGRAPH_ALIGNMENT::DISTRIBUTE);

    p2 = doc.addParagraph("Alignment3");
    p2->setAlignment(WD_PARAGRAPH_ALIGNMENT::JUSTIFY_HI);

    doc.addPicture("://139924.jpg");
    doc.addPicture("://139924.jpg", Inches::emus(1.25));
    doc.addPicture("://149341.jpg", Cm::emus(13), Cm::emus(10));
    doc.addPicture("://149341.jpg", Cm::emus(3), Cm::emus(10));
    QImage img("://185607.jpg");
    qDebug() << img.size();
    doc.addPicture(img, Cm::emus(5));
    doc.addPicture(img, Cm::emus(2));

    // Table
    Table *table = doc.addTable(3, 3);
    QList<Cell *> cells = table->rowCells(0);
    cells.at(0)->addText("Hello");
    cells.at(1)->addText("Word");
    cells.at(2)->addText("!!!");
    Paragraph *p1 = cells.at(2)->addParagraph();
    Run *r1 = p1->addRun();
    r1->addPicture(imagePath3, Cm::emus(3), Cm::emus(5));

    QList<Cell *> cells2 = table->rowCells(1);
    Cell *cell = cells2.at(0);
    Table *table2 = cell->addTable(5, 5, "MediumShading1");
    cells2 = table2->rowCells(1);
    cells2.at(0)->addText("Table!!!");

    p1 = cells2.at(2)->addParagraph();
    r1 = p1->addRun();

    r1->addPicture(imagePath2, Cm::emus(3), Cm::emus(5));

    table->addColumn();
    table->addRow();
    doc.save("aSave.docx");
}

void TestDocument::testImageInfo()
{

    QMimeDatabase base;
    QMimeType fileInfo = base.mimeTypeForFile(imagePath1);


    QImage image1(imagePath1);
    qDebug() << "Image Info" ;
    qDebug() << "Image name" << fileInfo.name() << "  suffixes " << fileInfo.preferredSuffix();
    qDebug() << fileInfo.preferredSuffix().toStdString().c_str();

    qDebug() << image1.rect() << "  " << image1.size();
    qDebug() << "dpiX" << image1.logicalDpiX() << "  dpiY  " << image1.logicalDpiY();
    //qDebug() << "preX" << image.dotsPerMeterX() << "  perY  " << image.dotsPerMeterY();
    qDebug() << "cacheKey1" << image1.cacheKey();

    QImage image11(imagePath1);
    QImage image2(imagePath2);
    qDebug() << "cacheKey11" << image11.cacheKey();
}

void TestDocument::testTable()
{
    Document doc(QStringLiteral("aSave.docx"));
    //Document doc;
    Table *table = doc.addTable(3, 3);
    QList<Cell *> cells = table->rowCells(0);
    cells.at(0)->addText("Hello");
    cells.at(1)->addText("Word");
    cells.at(2)->addText("!!!");
    Paragraph *p1 = cells.at(2)->addParagraph();
    Run *r1 = p1->addRun();
    r1->addPicture(imagePath3, Cm::emus(3), Cm::emus(10));

    qDebug() << QString::fromLatin1("----Cell index  ") <<  cells.at(1)->cellIndex()
             << QStringLiteral("--Row index  ") << cells.at(1)->rowIndex();

    QList<Cell *> cells2 = table->rowCells(1);
    Cell *cell = cells2.at(0);
    Table *table2 = cell->addTable(5, 5, "MediumShading1");
    cells2 = table2->rowCells(1);
    cells2.at(0)->addText("Table!!!");

    p1 = cells2.at(2)->addParagraph();
    r1 = p1->addRun();

    r1->addPicture(imagePath2, Cm::emus(3), Cm::emus(10));

    doc.addParagraph();

    table = doc.addTable(3, 3);
    table->setAlignment(WD_TABLE_ALIGNMENT::RIGHT);
    cells = table->rowCells(0);
    cells.at(0)->addText("Hello");

    table = doc.addTable(3, 3);
    table->setAlignment(WD_TABLE_ALIGNMENT::CENTER);
    cells = table->rowCells(0);
    cells.at(0)->addText("Hello");


    doc.save("atable.docx");
}

QTEST_APPLESS_MAIN(TestDocument)
#include "main.moc"

