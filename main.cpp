#include <QtTest>
#include <QString>
#include <QDebug>
#include <QByteArray>
#include <QFile>
#include <QFont>
#include <QTime>
#include <QColor>
#include <QDomDocument>

#include <docx/document.h>
#include <docx/text.h>


using namespace Docx;

class TestDocument : public QObject
{
    Q_OBJECT

public:
    TestDocument();

private Q_SLOTS:
    void testLoad();
};


TestDocument::TestDocument()
{

}

void TestDocument::testLoad()
{
    //demo.docx  Empty.docx default.docx
    Document doc(QStringLiteral("://default.docx"));
    doc.addHeading("MyTitle", 0);
    Paragraph *p = doc.addParagraph("helleWord");
    p->insertParagraphBefore("Before", "ListBullet");
    Run *addRun = p->addRun("AddRun");
    addRun->setStyle(QStringLiteral("Emphasis"));
    addRun->addText("Hello");

    doc.addParagraph();
    doc.addHeading("MyHead1");

    Paragraph *emptyP = doc.addParagraph();
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

    run = p3->addRun("Main");
    run->setAllcaps(true);

    run = p3->addRun("Main2");
    run->setDoubleStrike(true);

    run = p3->addRun("Main3");
    run->setItalic(true);

    doc.save("aSave.docx");
}


QTEST_APPLESS_MAIN(TestDocument)
#include "main.moc"
