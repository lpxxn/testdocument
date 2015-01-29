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
    Paragraph *p = doc.addParagraph("helleWord");
    p->insertParagraphBefore("Before", "ListBullet");
    Run *addRun = p->addRun("AddRun");
    addRun->setStyle(QStringLiteral("Emphasis"));
    addRun->addText("Hello");

    doc.addParagraph();
    Paragraph *emptyP = doc.addParagraph();
    emptyP->addRun("EmptyParagraph", "IntenseEmphasis");

    doc.addParagraph("West", "IntenseQuoteChar");

    Paragraph *p2 = doc.addParagraph("Two");
    Run *r2 = p2->addRun("Next");
    r2->setBold(true);
    r2->setBold(false);
    r2->setBold(true);

    doc.addParagraph("Heading1", "Heading1");
    doc.save("aSave.docx");
}


QTEST_APPLESS_MAIN(TestDocument)
#include "main.moc"
